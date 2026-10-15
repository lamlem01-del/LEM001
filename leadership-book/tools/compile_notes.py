"""Merge Phase 2a reading notes (phase2/notes/G*-events.jsonl) into data tables.

Usage:  python3 tools/compile_notes.py
Outputs: data/chronicle-events.csv   every event in text order (S##, ¶)
         data/case-evidence.csv      number of events and paragraph refs per case id
Prints invalid JSONL lines instead of silently dropping them.
"""
import csv
import glob
import json
import os
import re
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NOTES = os.path.join(ROOT, "phase2", "notes")
DATA = os.path.join(ROOT, "data")


def load():
    events, bad = [], []
    for path in sorted(glob.glob(os.path.join(NOTES, "G*-events.jsonl"))):
        group = os.path.basename(path).split("-")[0]
        for n, line in enumerate(open(path, encoding="utf-8"), 1):
            line = line.strip()
            if not line:
                continue
            try:
                e = json.loads(line)
            except json.JSONDecodeError as err:
                bad.append(f"{os.path.basename(path)}:{n}: {err}")
                continue
            e["group"] = group
            events.append(e)
    return events, bad


def key(e):
    m = re.match(r"S(\d+)", str(e.get("sid", "")))
    return (int(m.group(1)) if m else 999, int(e.get("para") or 0))


def main():
    events, bad = load()
    events.sort(key=key)
    with open(os.path.join(DATA, "chronicle-events.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ទីតាំង", "ប្រភេទ", "តួអង្គ", "L-id", "សង្ខេប", "កាលបរិច្ឆេទតាមអត្ថបទ", "គ.ស. (តាមអត្ថបទ)", "ករណី", "ក្រុមអាន"])
        for e in events:
            w.writerow([f"{e.get('sid')}¶{e.get('para')}", e.get("type", ""), e.get("actor", ""), e.get("l_id") or "",
                        e.get("summary_km", ""), e.get("date_text") or "", e.get("ce_year") or "",
                        "; ".join(e.get("cases") or []), e["group"]])
    by_case = defaultdict(list)
    for e in events:
        for c in e.get("cases") or []:
            by_case[c].append(f"{e.get('sid')}¶{e.get('para')}")
    with open(os.path.join(DATA, "case-evidence.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ករណី", "ចំនួនព្រឹត្តិការណ៍", "ទីតាំង"])
        for c in sorted(by_case, key=lambda x: (x[0], int(re.sub(r"\D", "", x) or 0))):
            w.writerow([c, len(by_case[c]), "; ".join(by_case[c])])
    print(f"{len(events)} events from {len({e['group'] for e in events})} groups; {len(by_case)} case ids")
    for b in bad:
        print("INVALID", b)


if __name__ == "__main__":
    main()
