#!/usr/bin/env python3
import sys
import zipfile
import json
import datetime
import re
import html
import xml.etree.ElementTree as ET
from pathlib import Path

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ALLOWED_DAYS = {"Pondělí", "Úterý", "Středa", "Čtvrtek", "Pátek", "Sobota"}

ORDER = [
    "1 rok", "1-1,5", "1,5-2", "2-2,5", "2,5-3",
    "I.z", "II.z", "III.z", "III.", "III. + 1 rok",
    "PŘS", "PŘS ", "PŘB"
]

LABELS = {
    "1 rok": "👶 1 rok",
    "1-1,5": "👶 1–1,5 roku",
    "1,5-2": "🫧 1,5–2 roky",
    "2-2,5": "🐳 2–2,5 roku",
    "2,5-3": "🏊 2,5–3 roky",
    "I.z": "I.z",
    "II.z": "II.z",
    "III.z": "III.z",
    "III.": "III.",
    "III. + 1 rok": "III. + 1 rok",
    "PŘS": "PŘS",
    "PŘS ": "PŘS",
    "PŘB": "PŘB",
}

SECTION_MAIN = {"1 rok", "1-1,5", "1,5-2", "2-2,5", "2,5-3"}


def col(ref):
    return re.match(r"[A-Z]+", ref).group(0)


def time_text(v):
    if v is None:
        return ""

    s = str(v).strip()

    try:
        x = float(s)
        mins = round((x % 1) * 24 * 60)
        return f"{mins // 60:02d}:{mins % 60:02d}"
    except Exception:
        pass

    m = re.match(r"^(\d{1,2}):(\d{2})", s)
    return f"{int(m.group(1)):02d}:{m.group(2)}" if m else s


def read_sheet(path, sheet_name):
    with zipfile.ZipFile(path) as z:
        shared = []

        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root:
                shared.append(
                    "".join((t.text or "") for t in si.iter(f"{{{M}}}t"))
                )

        wb = ET.fromstring(z.read("xl/workbook.xml"))

        rid = None
        for s in wb.find(f"{{{M}}}sheets"):
            if s.attrib.get("name") == sheet_name:
                rid = s.attrib[f"{{{R}}}id"]
                break

        if not rid:
            raise RuntimeError("WEB_DATA sheet not found")

        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))

        target = None
        for r in rels:
            if r.attrib.get("Id") == rid:
                target = r.attrib["Target"]
                break

        if not target:
            raise RuntimeError("WEB_DATA relationship not found")

        root = ET.fromstring(z.read("xl/" + target.lstrip("/")))

        rows = []

        for row in root.iter(f"{{{M}}}row"):
            vals = {}

            for c in row.findall(f"{{{M}}}c"):
                letter = col(c.attrib["r"])

                if letter not in {"A", "B", "C", "D", "E", "F", "G"}:
                    continue

                typ = c.attrib.get("t")
                v = c.find(f"{{{M}}}v")

                if typ == "inlineStr":
                    t = c.find(f".//{{{M}}}t")
                    val = t.text if t is not None else ""
                else:
                    val = v.text if v is not None else ""

                    if typ == "s" and val != "":
                        val = shared[int(val)]

                vals[letter] = val

            if vals:
                rows.append(vals)

        return rows


def status_class(status):
    if "Obsazeno" in status:
        return "full"
    if "Poslední" in status:
        return "last"
    return "free"


def build_html(payload):
    groups_html = []
    last_section = None

    for group in payload["groups"]:
        if group["section"] == "other" and last_section != "other":
            groups_html.append(
                '<div class="divider">Další skupiny</div>'
            )

        last_section = group["section"]

        lessons_html = []

        for lesson in group["lessons"]:
            lessons_html.append(
                '<div class="row">'
                f'<div>{html.escape(lesson["day"])}</div>'
                f'<div>{html.escape(lesson["time"])}</div>'
                f'<div class="status {status_class(lesson["status"])}">'
                f'{html.escape(lesson["status"])}</div>'
                '</div>'
            )

        groups_html.append(
            '<section class="group">'
            f'<div class="gtitle">{html.escape(group["label"])}</div>'
            '<div class="th">'
            '<div>Den</div><div>Čas</div><div>Dostupnost</div>'
            '</div>'
            + "".join(lessons_html)
            + '</section>'
        )

    generated = datetime.datetime.fromisoformat(
        payload["generated_at"]
    ).astimezone(
        datetime.timezone(datetime.timedelta(hours=2))
    )

    updated_text = generated.strftime("%d. %m. %Y %H:%M")

    return f"""<!doctype html>
<html lang="cs">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">

<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">

<title>Volná místa v kurzech plavání</title>

<style>
:root {{
  --blue:#95d9f1;
  --ink:#2f5966;
  --line:#d7edf5;
  --soft:#eef9fd;
}}

* {{
  box-sizing:border-box;
}}

html, body {{
  margin:0;
  background:transparent;
  font-family:Arial,Helvetica,sans-serif;
  color:#334;
}}

.card {{
  width:100%;
  max-width:900px;
  margin:0 auto;
  background:#fff;
  border:1px solid var(--line);
  border-radius:22px;
  overflow:hidden;
  box-shadow:0 6px 20px rgba(48,89,102,.13);
}}

.head {{
  text-align:center;
  background:var(--blue);
  color:var(--ink);
  padding:20px 22px 17px;
}}

.title {{
  font-size:28px;
  font-weight:700;
  line-height:1.15;
}}

.sub {{
  font-size:14px;
  margin-top:6px;
}}

.age {{
  font-size:12px;
  line-height:1.45;
  margin-top:8px;
  color:#315b68;
}}

.content {{
  padding:15px 16px 18px;
}}

.divider {{
  text-align:center;
  font-size:14px;
  font-weight:700;
  letter-spacing:.04em;
  text-transform:uppercase;
  color:var(--ink);
  margin:24px 0 10px;
}}

.group {{
  border:1px solid #e3eff3;
  border-radius:14px;
  overflow:hidden;
  margin:0 0 12px;
}}

.gtitle {{
  background:var(--soft);
  padding:10px 14px;
  font-size:17px;
  font-weight:700;
  color:#315b68;
}}

.th, .row {{
  display:grid;
  grid-template-columns:1.05fr .7fr 1.55fr;
  gap:8px;
  align-items:center;
  padding:8px 14px;
}}

.th {{
  font-size:12px;
  font-weight:700;
  color:#708087;
  border-top:1px solid #edf4f6;
  border-bottom:1px solid #edf4f6;
  background:#fafcfd;
}}

.row {{
  font-size:14px;
  border-bottom:1px solid #f0f4f5;
}}

.row:last-child {{
  border-bottom:0;
}}

.status {{
  font-weight:600;
}}

.full {{
  color:#aa2637;
}}

.last {{
  color:#9c7010;
}}

.free {{
  color:#23733b;
}}

.foot {{
  text-align:center;
  color:#776b77;
  font-size:12px;
  padding:2px 14px 16px;
}}

.updated {{
  margin-top:4px;
  color:#87949a;
}}

@media(max-width:560px) {{
  .title {{
    font-size:24px;
  }}

  .content {{
    padding:12px 10px 14px;
  }}

  .th, .row {{
    grid-template-columns:1fr .62fr 1.55fr;
    padding-left:10px;
    padding-right:10px;
    gap:6px;
  }}

  .row {{
    font-size:13px;
  }}

  .gtitle {{
    font-size:16px;
  }}
}}
</style>
</head>

<body>

<div class="card">

<header class="head">
  <div class="title">Volná místa v kurzech plavání</div>
  <div class="sub">
    Dostupnost se aktualizuje dle aktuální obsazenosti a je informativní.
  </div>
  <div class="age">
    I.z ≈ 6–7 měsíců &nbsp; • &nbsp;
    II.z ≈ 8–9 měsíců &nbsp; • &nbsp;
    III.z ≈ 10–12 měsíců
    <br>
    <span style="opacity:.85">
      z = začátečníci, věkové rozdělení je orientační
    </span>
  </div>
</header>

<main class="content">
{"".join(groups_html)}
</main>

<footer class="foot">
  Aktuální stav volných míst.
  <div class="updated">
    Poslední synchronizace: {updated_text}
  </div>
</footer>

</div>

<script>
// Každé 3 minuty znovu načti celý dokument.
// Přidání unikátního parametru pomáhá obejít cache v Google Sites.
setTimeout(() => {{
  const url = new URL(window.location.href);
  url.searchParams.set("_refresh", Date.now());
  window.location.replace(url.toString());
}}, 180000);
</script>

</body>
</html>
"""


def main(src, out):
    rows = read_sheet(src, "WEB_DATA")
    lessons = []

    for r in rows:
        day = str(r.get("A", "")).strip()
        cat = str(r.get("C", "")).strip()
        status = str(r.get("G", "")).strip()

        if day not in ALLOWED_DAYS or cat not in LABELS:
            continue

        if not (
            status.startswith("🟢")
            or status.startswith("🟡")
            or status.startswith("🔴")
        ):
            continue

        lessons.append({
            "day": day,
            "time": time_text(r.get("B")),
            "cat": cat,
            "status": status,
        })

    if len(lessons) < 50:
        raise RuntimeError(
            f"Privacy/sanity guard: expected many WEB_DATA lessons, got {len(lessons)}"
        )

    groups = []
    seen = set()

    for cat in ORDER:
        canonical = "PŘS" if cat == "PŘS " else cat

        if canonical in seen:
            continue

        ls = [
            x for x in lessons
            if ("PŘS" if x["cat"] == "PŘS " else x["cat"]) == canonical
        ]

        if not ls:
            continue

        seen.add(canonical)

        groups.append({
            "label": LABELS[cat],
            "section": "main" if canonical in SECTION_MAIN else "other",
            "lessons": [
                {
                    "day": x["day"],
                    "time": x["time"],
                    "status": x["status"],
                }
                for x in ls
            ],
        })

    payload = {
        "generated_at": datetime.datetime.now(
            datetime.timezone.utc
        ).isoformat(),
        "groups": groups,
    }

    out_path = Path(out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # JSON necháváme také zachovaný pro kontrolu.
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # Zároveň vytvoříme kompletní statickou HTML stránku.
    html_path = out_path.parent / "index.html"
    html_path.write_text(
        build_html(payload),
        encoding="utf-8",
    )


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(
            "usage: build_public.py SOURCE.xlsx OUTPUT.json"
        )

    main(sys.argv[1], sys.argv[2])
