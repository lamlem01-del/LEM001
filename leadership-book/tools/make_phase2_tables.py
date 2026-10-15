"""Generate the data-driven sections of the Phase 2 report (phase2/report/*b-*.md).

Usage:  python3 tools/make_phase2_tables.py
Inputs: phase2/verification/batch-*.json, data/case-studies-v2.csv, data/short-cases-v2.csv,
        data/leader-roster-v2.csv, data/date-problems.csv, data/chronicle-events.csv
"""
import csv
import glob
import json
import os
from collections import Counter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
OUT = os.path.join(ROOT, "phase2", "report")
KM = str.maketrans("0123456789", "០១២៣៤៥៦៧៨៩")


def km(n):
    return f"{n:,}".translate(KM)


def rows(name):
    path = os.path.join(DATA, name)
    if not os.path.exists(path):
        return None, []
    r = list(csv.reader(open(path, encoding="utf-8-sig")))
    return r[0], r[1:]


def cell(s, limit=None):
    s = (s or "").replace("|", "/").replace("\n", " ").strip()
    return s if not limit or len(s) <= limit else s[: limit - 1] + "…"


def table(header, body):
    out = ["| " + " | ".join(header) + " |", "|" + "|".join("---" for _ in header) + "|"]
    out += ["| " + " | ".join(cell(c) for c in r) + " |" for r in body]
    return "\n".join(out) + "\n"


def sources():
    items = []
    for path in sorted(glob.glob(os.path.join(ROOT, "phase2", "verification", "batch-*.json"))):
        items += json.load(open(path, encoding="utf-8"))
    core = [i for i in items if i.get("verified_title")]
    core.sort(key=lambda i: (i["id"][0] != "U", i["id"]))
    body = [[i["id"], i["verified_title"], i.get("authors_editors", ""), i.get("year_first", ""),
             cell(i.get("publisher_or_journal", ""), 80), i.get("status", "")] for i in core]
    status = Counter(i.get("status", "") for i in items)
    text = ("## ៦.៤ តារាងប្រភពដែលបានផ្ទៀងផ្ទាត់ព័ត៌មានគន្ថនិទ្ទេស\n\n"
            f"ធាតុដែលបានពិនិត្យសរុប៖ {km(len(items))} "
            f"(បានផ្ទៀងផ្ទាត់ {km(status['បានផ្ទៀងផ្ទាត់'])}, ផ្នែកខ្លះ {km(status['ផ្ទៀងផ្ទាត់ផ្នែកខ្លះ'])}, "
            f"មិនទាន់រកឃើញ {km(status['មិនទាន់រកឃើញកំណត់ត្រាផ្ទៀងផ្ទាត់'])})។ "
            "ស្ថានភាពអានរបស់ធាតុទាំងអស់គឺ «បានផ្ទៀងផ្ទាត់តែព័ត៌មានគន្ថនិទ្ទេស»។ "
            "ព័ត៌មានលម្អិត (ISBN តំណ ការកែតម្រូវ) មាននៅ `data/u300-audit-v2.csv` និង `data/a-candidates-v2.csv`។\n\n"
            + table(["លេខ", "ចំណងជើង", "អ្នកនិពន្ធ/កែសម្រួល", "ឆ្នាំដំបូង", "អ្នកបោះពុម្ព/ទស្សនាវដ្តី", "ស្ថានភាព"], body))
    open(os.path.join(OUT, "06b-sources-table.md"), "w", encoding="utf-8").write(text)


def cases():
    h, b = rows("case-studies-v2.csv")
    if not h:
        return
    want = ["លេខ", "ចំណងជើងករណី", "ផ្នែក និង ¶ សំខាន់", "កម្រិតភស្តុតាងបឋម", "លទ្ធផល", "ជំពូកដាក់"]
    idx = [h.index(c) for c in want]
    body = [[r[i] for i in idx] for r in b]
    text = "## តារាងករណីស៊ីជម្រៅ កំណែ ២\n\n" + table(want, body)
    hs, bs = rows("short-cases-v2.csv")
    if hs:
        want_s = ["លេខ", "ករណី", "ផ្នែក និង ¶", "ជំពូក"]
        idx_s = [hs.index(c) for c in want_s]
        text += "\n## តារាងករណីខ្លី កំណែ ២\n\n" + table(want_s, [[r[i] for i in idx_s] for r in bs])
    open(os.path.join(OUT, "05b-cases-table.md"), "w", encoding="utf-8").write(text)


def roster():
    h, b = rows("leader-roster-v2.csv")
    if not h:
        return
    pick = [0, 1, 2, 4, 5, 9, 11]
    body = [[cell(r[i], 90) for i in pick if i < len(r)] for r in b]
    text = (f"## តារាងបញ្ជីអ្នកដឹកនាំ កំណែ ២ ({km(len(b))} ជួរ)\n\n"
            "តារាងពេញ (ឈ្មោះផ្សេងៗ កាលបរិច្ឆេទ ភាពមិនស៊ីគ្នា ទីតាំង) មាននៅ `data/leader-roster-v2.csv`។\n\n"
            + table([h[i] for i in pick if i < len(h)], body))
    open(os.path.join(OUT, "11b-roster-table.md"), "w", encoding="utf-8").write(text)


def stats():
    _, ev = rows("chronicle-events.csv")
    _, dp = rows("date-problems.csv")
    types = Counter(r[1] for r in ev)
    dtypes = Counter(r[1] for r in dp) if dp else Counter()
    text = ("## ១.៤ លទ្ធផលជាលេខ\n\n"
            f"- ព្រឹត្តិការណ៍ដែលកត់ត្រា៖ **{km(len(ev))}** (`data/chronicle-events.csv`)\n"
            + "".join(f"  - {t}៖ {km(n)}\n" for t, n in types.most_common())
            + (f"- បញ្ហាកាលបរិច្ឆេទ និងតួលេខដែលកត់ត្រា៖ **{km(len(dp))}** (`data/date-problems.csv`)\n"
               + "".join(f"  - {t}៖ {km(n)}\n" for t, n in dtypes.most_common()) if dp else ""))
    open(os.path.join(OUT, "01b-stats.md"), "w", encoding="utf-8").write(text)


def main():
    os.makedirs(OUT, exist_ok=True)
    sources()
    cases()
    roster()
    stats()
    print("written:", sorted(os.listdir(OUT)))


if __name__ == "__main__":
    main()
