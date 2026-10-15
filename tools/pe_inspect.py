"""Inspect Word's binaries without REA (Python standard library only, read-only).

Reproduces every binary-level finding in docs/REA_ANALYSIS.md: PE header and sections, the resource
tree (types, counts, sizes) with SHA-256 of each payload, the embedded TYPELIB (optionally extracted to
a file), the import table, the export table, and the strings that show how WINWORD.EXE loads WWLIB.DLL.

    python tools/pe_inspect.py                       # the default Word install
    python tools/pe_inspect.py C:\\path\\WWLIB.DLL     # any PE file
    python tools/pe_inspect.py --extract-typelib out.tlb
"""

from __future__ import annotations

import argparse
import hashlib
import os
import struct

OFFICE = r"C:\Program Files\Microsoft Office\root\Office16"
DEFAULT = ["WINWORD.EXE", "WWLIB.DLL", "MSWORD.OLB"]
RT = {1: "CURSOR", 2: "BITMAP", 3: "ICON", 4: "MENU", 5: "DIALOG", 6: "STRING", 9: "ACCELERATOR",
      10: "RCDATA", 11: "MESSAGETABLE", 12: "GROUP_CURSOR", 14: "GROUP_ICON", 16: "VERSION", 24: "MANIFEST"}
MACHINE = {0x8664: "x64", 0x14C: "x86", 0xAA64: "ARM64"}


class PE:
    def __init__(self, path: str):
        self.path = path
        self.data = open(path, "rb").read()
        d = self.data
        if d[:2] != b"MZ":
            raise ValueError("not a PE file")
        pe = struct.unpack_from("<I", d, 0x3C)[0]
        self.machine, self.nsec = struct.unpack_from("<HH", d, pe + 4)
        opt_size = struct.unpack_from("<H", d, pe + 20)[0]
        self.opt = pe + 24
        self.plus = struct.unpack_from("<H", d, self.opt)[0] == 0x20B
        dd = self.opt + (112 if self.plus else 96)
        self.dirs = [struct.unpack_from("<II", d, dd + 8 * i) for i in range(16)]
        sec = self.opt + opt_size
        self.sections = []
        for i in range(self.nsec):
            name = d[sec + i * 40: sec + i * 40 + 8].rstrip(b"\0").decode("latin-1")
            vsz, va, rsz, raw = struct.unpack_from("<IIII", d, sec + i * 40 + 8)
            self.sections.append((name, va, vsz, raw, rsz))

    def off(self, rva: int) -> int | None:
        for _, va, vsz, raw, rsz in self.sections:
            if va <= rva < va + max(vsz, rsz):
                return rva - va + raw
        return None

    def cstr(self, off: int) -> str:
        return self.data[off:self.data.index(b"\0", off)].decode("latin-1")

    # --- tables ------------------------------------------------------------------------
    def exports(self) -> list[str]:
        rva, size = self.dirs[0]
        if not rva:
            return []
        o = self.off(rva)
        n_names = struct.unpack_from("<I", self.data, o + 24)[0]
        names = self.off(struct.unpack_from("<I", self.data, o + 32)[0])
        return [self.cstr(self.off(struct.unpack_from("<I", self.data, names + 4 * k)[0])) for k in range(n_names)]

    def imports(self, per_dll: int = 8) -> list[tuple[str, list[str]]]:
        rva, _ = self.dirs[1]
        if not rva:
            return []
        out, o = [], self.off(rva)
        width = 8 if self.plus else 4
        flag = 1 << (63 if self.plus else 31)
        while True:
            ilt, _, _, name_rva, iat = struct.unpack_from("<IIIII", self.data, o)
            if not name_rva:
                break
            fns, t = [], self.off(ilt or iat)
            while t is not None and len(fns) < per_dll:
                v = struct.unpack_from("<Q" if self.plus else "<I", self.data, t)[0]
                if not v:
                    break
                fns.append(f"#{v & 0xFFFF}" if v & flag else self.cstr(self.off(v & 0x7FFFFFFF) + 2))
                t += width
            out.append((self.cstr(self.off(name_rva)), fns))
            o += 20
        return out

    def resources(self) -> list[dict]:
        rva, _ = self.dirs[2]
        if not rva:
            return []
        base = self.off(rva)
        out = []

        def name_of(entry_name):
            if entry_name & 0x80000000:
                p = base + (entry_name & 0x7FFFFFFF)
                n = struct.unpack_from("<H", self.data, p)[0]
                return self.data[p + 2:p + 2 + 2 * n].decode("utf-16-le")
            return entry_name

        def walk(off, path):
            named, ids = struct.unpack_from("<HH", self.data, off + 12)
            for i in range(named + ids):
                ename, target = struct.unpack_from("<II", self.data, off + 16 + 8 * i)
                key = name_of(ename)
                if target & 0x80000000:
                    walk(base + (target & 0x7FFFFFFF), path + [key])
                else:
                    drva, size = struct.unpack_from("<II", self.data, base + target)
                    foff = self.off(drva)
                    payload = self.data[foff:foff + size]
                    typ = path[0] if path else key
                    out.append({"type": RT.get(typ, typ) if isinstance(typ, int) else typ,
                                "name": path[1] if len(path) > 1 else None, "lang": key,
                                "offset": foff, "size": size, "sha256": hashlib.sha256(payload).hexdigest()})
        walk(base, [])
        return out


def report(path: str, extract_typelib: str | None) -> None:
    pe = PE(path)
    size = os.path.getsize(path)
    print(f"\n{path}\n  {size:,} bytes, sha256 {hashlib.sha256(pe.data).hexdigest()}")
    print(f"  {'PE32+' if pe.plus else 'PE32'} {MACHINE.get(pe.machine, hex(pe.machine))}, sections: "
          + ", ".join(f"{n} ({rsz:,})" for n, _, _, _, rsz in pe.sections))
    res = pe.resources()
    if res:
        summary: dict = {}
        for r in res:
            k = str(r["type"])
            summary.setdefault(k, [0, 0])
            summary[k][0] += 1
            summary[k][1] += r["size"]
        print(f"  resources: {len(res)} - " + ", ".join(f"{k} x{c} ({b:,} bytes)" for k, (c, b) in summary.items()))
        for r in res:
            if r["type"] == "TYPELIB":
                print(f"  TYPELIB name={r['name']} lang={r['lang']} at file offset {r['offset']:,}, "
                      f"{r['size']:,} bytes, sha256 {r['sha256']}")
                if extract_typelib:
                    with open(extract_typelib, "wb") as fh:
                        fh.write(pe.data[r["offset"]:r["offset"] + r["size"]])
                    print(f"    extracted to {extract_typelib}")
    exp = pe.exports()
    if exp:
        plain = [e for e in exp if not e.startswith("?")]
        print(f"  exports: {len(exp)} ({len(exp) - len(plain)} C++ mangled); plain: {plain}")
    imp = pe.imports()
    if imp:
        print(f"  imports from {len(imp)} DLLs:")
        for dll, fns in imp:
            print(f"    {dll}: {', '.join(fns)}")
    for s in ("wwlib", "FMain", "LoadLibrary", "GetProcAddress", "DllGetClassObject"):
        a, u = pe.data.count(s.encode()), pe.data.count(s.encode("utf-16-le"))
        if a or u:
            print(f"  string {s!r}: ascii x{a}, utf-16 x{u}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*")
    ap.add_argument("--extract-typelib", help="write MSWORD.OLB's TYPELIB resource to this file")
    a = ap.parse_args()
    files = a.files or [os.path.join(OFFICE, f) for f in DEFAULT]
    for f in files:
        if os.path.exists(f):
            report(f, a.extract_typelib if f.upper().endswith(".OLB") else None)
        else:
            print(f"\n{f}: not found")


if __name__ == "__main__":
    main()
