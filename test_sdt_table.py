"""Offline regression: a table cell wrapped in a content control (w:sdt) must be counted.

Cover pages bind document properties (version, status) to table cells through w:sdt, and the
parse used to skip those cells, so every paragraph number after them was off by the missing cells.
No Word needed.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from docx_context import Context  # noqa: E402

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def cell(text):
    return f"<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"


def package(body):
    return (
        '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
        '<pkg:part pkg:name="/word/document.xml">'
        f'<pkg:xmlData><w:document xmlns:w="{W_NS}"><w:body>{body}</w:body></w:document></pkg:xmlData>'
        "</pkg:part></pkg:package>"
    )


failures = []


def check(cond, what):
    print(("  ok   " if cond else "  FAIL ") + what)
    if not cond:
        failures.append(what)


plain = f"<w:tbl><w:tr>{cell('Version')}{cell('0.1')}</w:tr></w:tbl><w:p><w:r><w:t>After.</w:t></w:r></w:p>"
wrapped = (f"<w:tbl><w:tr>{cell('Version')}"
           f"<w:sdt><w:sdtPr/><w:sdtContent>{cell('0.1')}</w:sdtContent></w:sdt></w:tr></w:tbl>"
           "<w:p><w:r><w:t>After.</w:t></w:r></w:p>")
row_wrapped = (f"<w:tbl><w:sdt><w:sdtPr/><w:sdtContent><w:tr>{cell('Version')}{cell('0.1')}</w:tr>"
               "</w:sdtContent></w:sdt></w:tbl><w:p><w:r><w:t>After.</w:t></w:r></w:p>")

for name, body in (("plain", plain), ("cell in content control", wrapped), ("row in content control", row_wrapped)):
    ctx = Context(package(body))
    texts = [m["plain"] for m in ctx.para_meta]
    check(len(ctx.para_meta) == 4, f"{name}: 2 cells + row end + 1 body paragraph = 4 (got {len(ctx.para_meta)})")
    check("0.1" in texts, f"{name}: the cell text 0.1 is read")
    check(texts[-1] == "After.", f"{name}: 'After.' is paragraph 4")

print("FAILURES:" if failures else "ALL CHECKS PASSED", *failures, sep="\n  ")
sys.exit(1 if failures else 0)
