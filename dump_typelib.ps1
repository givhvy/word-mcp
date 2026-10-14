# Decodes Word's COM type library (MSWORD.OLB) into word-api.json for word_api_search.
# Both parameters are optional: the type library is found through the registry.
param([string]$Olb, [string]$OutJson = (Join-Path $PSScriptRoot "word-api.json"))

if (-not $Olb) {
    $key = "Registry::HKEY_CLASSES_ROOT\TypeLib\{00020905-0000-0000-C000-000000000046}"
    $ver = Get-ChildItem $key -ErrorAction Stop | Sort-Object { [version]($_.PSChildName -replace '[^0-9.]', '') } | Select-Object -Last 1
    foreach ($arch in "win64", "win32") {
        $p = (Get-ItemProperty "$($ver.PSPath)\0\$arch" -ErrorAction SilentlyContinue).'(default)'
        if ($p -and (Test-Path $p)) { $Olb = $p; break }
    }
    if (-not $Olb) { throw "Word type library not found; is Microsoft Word (desktop) installed?" }
}

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using DllImport = System.Runtime.InteropServices.DllImportAttribute;
using CharSet = System.Runtime.InteropServices.CharSet;
using Marshal = System.Runtime.InteropServices.Marshal;
using VarEnum = System.Runtime.InteropServices.VarEnum;
using System.Text;

public static class TlbDump
{
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern void LoadTypeLibEx(string file, int regKind, out ITypeLib lib);

    static string Vt(ITypeInfo ti, TYPEDESC td)
    {
        VarEnum vt = (VarEnum)td.vt;
        if (vt == VarEnum.VT_PTR || vt == VarEnum.VT_SAFEARRAY)
        {
            TYPEDESC inner = (TYPEDESC)Marshal.PtrToStructure(td.lpValue, typeof(TYPEDESC));
            return Vt(ti, inner) + (vt == VarEnum.VT_PTR ? "*" : "[]");
        }
        if (vt == VarEnum.VT_USERDEFINED)
        {
            ITypeInfo r; ti.GetRefTypeInfo(unchecked((int)(td.lpValue.ToInt64() & 0xFFFFFFFF)), out r);
            string n, d, h; int c; r.GetDocumentation(-1, out n, out d, out c, out h);
            return n;
        }
        return vt.ToString().Replace("VT_", "").ToLower();
    }

    static string Esc(string s)
    {
        if (s == null) return "";
        var b = new StringBuilder();
        foreach (char ch in s)
        {
            if (ch == '"' || ch == '\\') b.Append('\\').Append(ch);
            else if (ch < 32) b.Append(' ');
            else b.Append(ch);
        }
        return b.ToString();
    }

    public static string Dump(string path)
    {
        ITypeLib lib; LoadTypeLibEx(path, 2, out lib);
        var sb = new StringBuilder("[");
        int count = lib.GetTypeInfoCount();
        for (int i = 0; i < count; i++)
        {
            ITypeInfo ti; lib.GetTypeInfo(i, out ti);
            string name, doc, help; int ctx; ti.GetDocumentation(-1, out name, out doc, out ctx, out help);
            IntPtr pa; ti.GetTypeAttr(out pa);
            TYPEATTR ta = (TYPEATTR)Marshal.PtrToStructure(pa, typeof(TYPEATTR));
            if (i > 0) sb.Append(",");
            sb.AppendFormat("{{\"name\":\"{0}\",\"kind\":\"{1}\",\"guid\":\"{2}\",\"doc\":\"{3}\",\"members\":[",
                Esc(name), ta.typekind.ToString().Replace("TKIND_", "").ToLower(), ta.guid, Esc(doc));
            bool first = true;
            for (int f = 0; f < ta.cFuncs; f++)
            {
                IntPtr pf; ti.GetFuncDesc(f, out pf);
                FUNCDESC fd = (FUNCDESC)Marshal.PtrToStructure(pf, typeof(FUNCDESC));
                string[] names = new string[fd.cParams + 1]; int got;
                ti.GetNames(fd.memid, names, names.Length, out got);
                var ps = new List<string>();
                for (int p = 0; p < fd.cParams; p++)
                {
                    IntPtr ep = new IntPtr(fd.lprgelemdescParam.ToInt64() + p * Marshal.SizeOf(typeof(ELEMDESC)));
                    ELEMDESC ed = (ELEMDESC)Marshal.PtrToStructure(ep, typeof(ELEMDESC));
                    bool opt = (ed.desc.paramdesc.wParamFlags & PARAMFLAG.PARAMFLAG_FOPT) != 0;
                    string pn = p + 1 < got ? names[p + 1] : ("arg" + p);
                    ps.Add(string.Format("{{\"name\":\"{0}\",\"type\":\"{1}\",\"optional\":{2}}}",
                        Esc(pn), Esc(Vt(ti, ed.tdesc)), opt ? "true" : "false"));
                }
                string kind = fd.invkind == INVOKEKIND.INVOKE_FUNC ? "method"
                    : fd.invkind == INVOKEKIND.INVOKE_PROPERTYGET ? "get"
                    : fd.invkind == INVOKEKIND.INVOKE_PROPERTYPUT ? "put" : "putref";
                string md, mh, mn; int mc; ti.GetDocumentation(fd.memid, out mn, out md, out mc, out mh);
                if (!first) sb.Append(","); first = false;
                sb.AppendFormat("{{\"name\":\"{0}\",\"kind\":\"{1}\",\"dispid\":{2},\"returns\":\"{3}\",\"hidden\":{4},\"params\":[{5}]}}",
                    Esc(names[0]), kind, fd.memid, Esc(Vt(ti, fd.elemdescFunc.tdesc)),
                    (fd.wFuncFlags & 0x40) != 0 ? "true" : "false", string.Join(",", ps));
                ti.ReleaseFuncDesc(pf);
            }
            for (int v = 0; v < ta.cVars; v++)
            {
                IntPtr pv; ti.GetVarDesc(v, out pv);
                VARDESC vd = (VARDESC)Marshal.PtrToStructure(pv, typeof(VARDESC));
                string[] vn = new string[1]; int got; ti.GetNames(vd.memid, vn, 1, out got);
                string val = "";
                if (vd.varkind == VARKIND.VAR_CONST)
                    val = Convert.ToString(Marshal.GetObjectForNativeVariant(vd.desc.lpvarValue));
                if (!first) sb.Append(","); first = false;
                sb.AppendFormat("{{\"name\":\"{0}\",\"kind\":\"const\",\"value\":\"{1}\"}}", Esc(vn[0]), Esc(val));
                ti.ReleaseVarDesc(pv);
            }
            sb.Append("]}");
            ti.ReleaseTypeAttr(pa);
        }
        return sb.Append("]").ToString();
    }
}
'@

[IO.File]::WriteAllText($OutJson, [TlbDump]::Dump($Olb), (New-Object Text.UTF8Encoding $false))
"wrote $OutJson"
