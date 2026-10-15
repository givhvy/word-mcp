"""Does a user reading / scrolling the document disturb an agent writing to it?

A simulated user pages up and down in the Word window non-stop (cua-driver PageDown/PageUp keystrokes,
delivered in the background, which also moves Word's cursor), while the agent runs a full workload
through the MCP: section_end inserts by heading, formatting, comments, reads, section rewrites, page
renders. The same workload also runs without the user, as a baseline. Checks that every edit lands in
the right section, that scrolling is not reported as an edit, and how latency changes.

Requires cua-driver (https://cua.ai) running in the interactive session."""

import asyncio
import json
import os
import statistics
import subprocess
import sys
import threading
import time

import pythoncom
import win32com.client
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # python -I leaves this folder off sys.path
import tests_common  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "test-output")
SECTIONS = 30
ROUNDS = 8
failures: list[str] = []


def check(cond, what):
    print(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        failures.append(what)


def cua(tool, args, timeout=30):
    r = subprocess.run(["cua-driver", tool], input=json.dumps(args), capture_output=True, text=True, timeout=timeout)
    return r.returncode, (r.stdout + r.stderr).strip()


def word_pid() -> int:
    out = subprocess.check_output(["powershell", "-NoProfile", "-Command",
                                   "(Get-Process WINWORD | Select-Object -First 1).Id"], text=True)
    return int(out.strip())


class User(threading.Thread):
    """Pages through the document like a reader: PageDown x4, PageUp x4, ... until stopped."""

    def __init__(self, pid, window_id, doc_name):
        super().__init__(daemon=True)
        self.pid, self.window_id, self.doc_name = pid, window_id, doc_name
        self.stop = threading.Event()
        self.keys = 0
        self.failed = 0
        self.positions: list[float] = []

    def run(self):
        pythoncom.CoInitialize()
        app = win32com.client.dynamic.Dispatch(win32com.client.GetActiveObject("Word.Application")._oleobj_)
        pane = app.Documents(self.doc_name).Windows(1).ActivePane
        k = 0
        while not self.stop.is_set():
            key = "pagedown" if (k // 4) % 2 == 0 else "pageup"
            code, _ = cua("press_key", {"pid": self.pid, "window_id": self.window_id, "key": key})
            self.keys += 1
            self.failed += code != 0
            try:
                self.positions.append(pane.VerticalPercentScrolled)
            except pythoncom.com_error:
                pass
            k += 1


async def call(s, tool, timings, label=None, **args):
    t = time.perf_counter()
    r = await s.call_tool(tool, args)
    timings.append((label or tool, (time.perf_counter() - t) * 1000))
    text = "".join(c.text for c in r.content if c.type == "text")
    if r.isError:
        raise AssertionError(f"{tool}: {text}")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


async def make_doc(s, timings):
    doc = (await call(s, "word_new_document", timings, text="Concurrency Test"))["name"]
    md = "\n\n".join(f"## Module {k}\n\nModule {k} handles one part of the pipeline and has a few paragraphs of text "
                     f"so that the document is long enough to scroll through.\n\n"
                     f"- Responsibility A of module {k}\n- Responsibility B of module {k}\n\n"
                     f"Closing paragraph of module {k}." for k in range(1, SECTIONS + 1))
    await call(s, "word_insert_markdown", timings, markdown=md, document=doc)
    return doc


async def workload(s, doc, timings, user=None):
    """The agent's edits; returns the markers it wrote, keyed by section."""
    await call(s, "word_context", timings, document=doc)
    ch = await call(s, "word_changes", timings, document=doc)
    quiet = [ch]
    for k in range(1, ROUNDS + 1):
        sec = 3 * k  # spread over the document: modules 3, 6, ..., 24
        marker = f"AGENT-NOTE-{k}"
        await call(s, "word_insert_markdown", timings, label="insert section_end",
                   markdown=f"{marker}: notes for module {sec}.\n\n| key | value |\n|---|---|\n| owner | team {k} |",
                   position="section_end", anchor=f"Module {sec}", document=doc)
        await call(s, "word_format", timings, find=marker, bold=True, document=doc)
        await call(s, "word_add_comment", timings, comment=f"check {marker}", find=marker, document=doc)
        await call(s, "word_context", timings, label="word_context", stats=False, document=doc)
        if k % 4 == 0:
            await call(s, "word_replace_section", timings, label="replace section", anchor=f"Module {sec + 1}",
                       markdown=f"REWRITTEN-{sec + 1}: this section was rewritten by the agent.", document=doc)
        await call(s, "word_page_image", timings, label="page image", pages="1", document=doc)
        quiet.append(await call(s, "word_changes", timings, document=doc))
    return quiet


def verify(s_texts: list[str], label: str):
    """Every marker sits inside its own section, once; rewritten sections hold only the new text."""
    heads = {t: i for i, t in enumerate(s_texts) if t.startswith("Module ") and len(t) <= 10}
    ok_place, ok_once = True, True
    for k in range(1, ROUNDS + 1):
        sec = 3 * k
        marker = f"AGENT-NOTE-{k}"
        hits = [i for i, t in enumerate(s_texts) if t.startswith(marker)]
        ok_once &= len(hits) == 1
        start = heads.get(f"Module {sec}", -1)
        end = heads.get(f"Module {sec + 1}", len(s_texts))
        if not hits or not start < hits[0] < end:
            ok_place = False
            print(f"    {marker}: at {hits}, section Module {sec} spans {start}..{end}")
    check(ok_once, f"{label}: each of the {ROUNDS} agent notes written exactly once")
    check(ok_place, f"{label}: every note landed inside its own section (by heading anchor)")
    rewritten = [sec + 1 for sec in (12, 24)]
    ok_rw = True
    for sec in rewritten:
        i = heads.get(f"Module {sec}")
        nxt = heads.get(f"Module {sec + 1}", len(s_texts))
        body = [t for t in s_texts[i + 1:nxt] if t] if i is not None else []
        ok_rw &= body == [f"REWRITTEN-{sec}: this section was rewritten by the agent."]
    check(ok_rw, f"{label}: rewritten sections contain only the new text")


async def main():
    tests_common.setup()
    os.makedirs(OUT, exist_ok=True)
    pid = word_pid()
    pythoncom.CoInitialize()
    app = win32com.client.dynamic.Dispatch(win32com.client.GetActiveObject("Word.Application")._oleobj_)
    params = StdioServerParameters(command=sys.executable, args=["-I", os.path.join(HERE, "word_mcp_server.py")])
    async with stdio_client(params) as (r, w), ClientSession(r, w) as s:
        await s.initialize()
        before = {d["name"] for d in json.loads((await s.call_tool("word_status", {})).content[0].text)["documents"]}
        results = {}
        try:
            for mode in ("baseline", "user scrolling"):
                print(f"\n== {mode}")
                timings: list[tuple[str, float]] = []
                doc = await make_doc(s, timings)
                timings.clear()
                user = None
                if mode == "user scrolling":
                    wins = json.loads(cua("list_windows", {"pid": pid})[1])["windows"]
                    win = next(x for x in wins if x["title"].startswith(doc + " "))
                    user = User(pid, win["window_id"], doc)
                    user.start()
                    time.sleep(2.0)  # the user is already reading when the agent starts
                t0 = time.perf_counter()
                quiet = await workload(s, doc, timings, user)
                elapsed = time.perf_counter() - t0
                if user:
                    cua("get_window_state", {"pid": pid, "window_id": user.window_id,
                                             "screenshot_out_file": os.path.join(OUT, "concurrent-user-view.png")})
                    user.stop.set()
                    user.join(timeout=10)
                texts = [p["text"] for p in (await call(s, "word_read", [], document=doc))["paragraphs"]]
                verify(texts, mode)
                comments = (await call(s, "word_list_comments", [], document=doc))["comment_count"]
                check(comments == ROUNDS, f"{mode}: {ROUNDS} comments added ({comments})")
                changed = [q for q in quiet[1:] if q.get("changed")]
                check(not changed, f"{mode}: word_changes never reports scrolling/cursor moves as edits"
                      + (f" ({changed[0]})" if changed else ""))
                if user:
                    distinct = len({round(p) for p in user.positions})
                    print(f"  user: {user.keys} PageUp/PageDown presses, {user.failed} failed, "
                          f"view positions seen: {distinct} distinct ({min(user.positions):.0f}%..{max(user.positions):.0f}%)")
                    check(user.keys >= 20 and distinct >= 5, "the simulated user really was scrolling the whole time")
                    check(user.failed == 0, "user input never blocked")
                med = statistics.median(ms for _, ms in timings)
                results[mode] = (elapsed, med, timings)
                print(f"  workload {elapsed:.1f} s, {len(timings)} calls, median {med:.0f} ms")
        finally:
            for name in [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]:
                if name not in before:
                    app.Documents(name).Close(0)
    (e0, m0, t0_), (e1, m1, t1_) = results["baseline"], results["user scrolling"]
    print(f"\nbaseline {e0:.1f} s (median {m0:.0f} ms)  vs  user scrolling {e1:.1f} s (median {m1:.0f} ms)")
    by = {}
    for label, ms in t0_:
        by.setdefault(label, [[], []])[0].append(ms)
    for label, ms in t1_:
        by.setdefault(label, [[], []])[1].append(ms)
    for label, (a, b) in by.items():
        if a and b:
            print(f"  {label:20s} {statistics.median(a):7.0f} ms -> {statistics.median(b):7.0f} ms")
    check(e1 < e0 * 2.5, "scrolling slows the agent by less than 2.5x overall")
    print("FAILURES:" if failures else "ALL CHECKS PASSED", *failures, sep="\n  ")
    sys.exit(1 if failures else 0)


asyncio.run(main())
