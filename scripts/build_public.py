#!/usr/bin/env python3
import sys, zipfile, json, datetime, re, xml.etree.ElementTree as ET
from pathlib import Path

M="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R="http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ALLOWED_DAYS={"Pondělí","Úterý","Středa","Čtvrtek","Pátek","Sobota"}
ORDER=["1 rok","1-1,5","1,5-2","2-2,5","2,5-3","I.z","II.z","III.z","III.","III. + 1 rok","PŘS","PŘS ","PŘB"]
LABELS={
 "1 rok":"👶 1 rok","1-1,5":"👶 1–1,5 roku","1,5-2":"🫧 1,5–2 roky","2-2,5":"🐳 2–2,5 roku",
 "2,5-3":"🏊 2,5–3 roky","I.z":"I.z","II.z":"II.z","III.z":"III.z","III.":"III.","III. + 1 rok":"III. + 1 rok",
 "PŘS":"PŘS","PŘS ":"PŘS","PŘB":"PŘB"
}
SECTION_MAIN={"1 rok","1-1,5","1,5-2","2-2,5","2,5-3"}

def col(ref): return re.match(r"[A-Z]+",ref).group(0)
def time_text(v):
    if v is None: return ""
    s=str(v).strip()
    try:
        x=float(s)
        mins=round((x%1)*24*60)
        return f"{mins//60:02d}:{mins%60:02d}"
    except: pass
    m=re.match(r"^(\d{1,2}):(\d{2})",s)
    return f"{int(m.group(1)):02d}:{m.group(2)}" if m else s

def read_sheet(path, sheet_name):
    with zipfile.ZipFile(path) as z:
        shared=[]
        if "xl/sharedStrings.xml" in z.namelist():
            root=ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root:
                shared.append("".join((t.text or "") for t in si.iter(f"{{{M}}}t")))
        wb=ET.fromstring(z.read("xl/workbook.xml"))
        rid=None
        for s in wb.find(f"{{{M}}}sheets"):
            if s.attrib.get("name")==sheet_name:
                rid=s.attrib[f"{{{R}}}id"]; break
        if not rid: raise RuntimeError("WEB_DATA sheet not found")
        rels=ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target=None
        for r in rels:
            if r.attrib.get("Id")==rid: target=r.attrib["Target"]; break
        if not target: raise RuntimeError("WEB_DATA relationship not found")
        root=ET.fromstring(z.read("xl/"+target.lstrip("/")))
        rows=[]
        for row in root.iter(f"{{{M}}}row"):
            vals={}
            for c in row.findall(f"{{{M}}}c"):
                letter=col(c.attrib["r"])
                if letter not in {"A","B","C","D","E","F","G"}: continue
                typ=c.attrib.get("t"); v=c.find(f"{{{M}}}v")
                if typ=="inlineStr":
                    t=c.find(f".//{{{M}}}t"); val=t.text if t is not None else ""
                else:
                    val=v.text if v is not None else ""
                    if typ=="s" and val!="": val=shared[int(val)]
                vals[letter]=val
            if vals: rows.append(vals)
        return rows

def main(src, out):
    rows=read_sheet(src,"WEB_DATA")
    lessons=[]
    for r in rows:
        day=str(r.get("A","")).strip()
        cat=str(r.get("C","")).strip()
        status=str(r.get("G","")).strip()
        if day not in ALLOWED_DAYS or cat not in LABELS: continue
        if not (status.startswith("🟢") or status.startswith("🟡") or status.startswith("🔴")): continue
        lessons.append({"day":day,"time":time_text(r.get("B")),"cat":cat,"status":status})
    if len(lessons)<50:
        raise RuntimeError(f"Privacy/sanity guard: expected many WEB_DATA lessons, got {len(lessons)}")
    groups=[]
    seen=set()
    for cat in ORDER:
        canonical="PŘS" if cat=="PŘS " else cat
        if canonical in seen: continue
        ls=[x for x in lessons if ("PŘS" if x["cat"]=="PŘS " else x["cat"])==canonical]
        if not ls: continue
        seen.add(canonical)
        groups.append({"label":LABELS[cat],"section":"main" if canonical in SECTION_MAIN else "other",
                       "lessons":[{"day":x["day"],"time":x["time"],"status":x["status"]} for x in ls]})
    payload={"generated_at":datetime.datetime.now(datetime.timezone.utc).isoformat(),"groups":groups}
    Path(out).parent.mkdir(parents=True,exist_ok=True)
    Path(out).write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8")

if __name__=="__main__":
    if len(sys.argv)!=3: raise SystemExit("usage: build_public.py SOURCE.xlsx OUTPUT.json")
    main(sys.argv[1],sys.argv[2])
