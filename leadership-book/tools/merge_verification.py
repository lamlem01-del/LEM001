"""Merge Phase 2b verification results into version-2 source tables.

Usage:  python3 tools/merge_verification.py
Inputs: data/u300-audit.csv, data/a-candidates.csv (Phase 1 snapshots, left unchanged)
        phase2/verification/batch-*.json
Outputs: data/u300-audit-v2.csv, data/a-candidates-v2.csv

All verification in Phase 2b was done from search-result snippets and catalogue records;
no full text was read. The "វិធីផ្ទៀងផ្ទាត់" column says so for every verified row.
"""
import csv
import glob
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
METHOD = "លទ្ធផលស្វែងរក និងកំណត់ត្រាកាតាឡុក (មិនទាន់បើកទំព័រពេញ)"

# Candidates found during Phase 2 (not in the Phase 1 A-list). Bibliographic details
# come from the verification agents' snippets and remain to be confirmed on the page.
NEW_A = [
    ["A024", "Theara Thun", "Treason and loyalty to the royal court: The Kan narrative in nineteenth- and early-twentieth-century Khmer chronicle manuscripts (Journal of the Royal Asiatic Society, ser. 3, 31(2), 2021, pp. 343–362)",
     "ទស្សនៈផ្ទុយសំខាន់សម្រាប់ C09៖ សេចក្តីសង្ខេបអះអាងថារឿងស្ដេចកនត្រូវបានសាងសង់ដោយអ្នកសរសេរពង្សាវតារដើមសតវត្សទី២០", "ជ.២; ជ.៧; ជ.៩ ផ្នែក៨",
     "ផ្ទៀងផ្ទាត់ផ្នែកខ្លះ", "មិនទាន់អាន (ឃើញតែសេចក្តីសង្ខេបក្នុងលទ្ធផលស្វែងរក)"],
    ["A025", "Thibodi Buakamsri", "និក្ខេបបទអនុបណ្ឌិត សាកលវិទ្យាល័យជូឡាឡុងកន ស្តីពី «ឯកសារមហាបុរសខ្មែរ» (ចំណងជើងពិតត្រូវរក)",
     "ការសិក្សាដោយផ្ទាល់លើឯកសារគោល — សម្រាប់ជំពូក ២", "ជ.២", "មិនទាន់ផ្ទៀងផ្ទាត់", "មិនទាន់អាន"],
]


def load_verification():
    out = {}
    for path in sorted(glob.glob(os.path.join(ROOT, "phase2", "verification", "batch-*.json"))):
        for item in json.load(open(path, encoding="utf-8")):
            out[item["id"]] = item
    return out


def corrected(v):
    parts = [v.get("verified_title"), v.get("authors_editors"), v.get("year_first") and f"បោះពុម្ពលើកដំបូង {v['year_first']}",
             v.get("publisher_or_journal"), v.get("isbn_or_doi"), v.get("language")]
    return " · ".join(p for p in parts if p)


def main():
    ver = load_verification()
    rows = list(csv.reader(open(os.path.join(DATA, "u300-audit.csv"), encoding="utf-8-sig")))
    header, body = rows[0], rows[1:]
    header = header + ["វិធីផ្ទៀងផ្ទាត់", "សេចក្តីពិពណ៌នាបឋម (ពីអ្វីដែលបានឃើញ)"]
    idx = {h: i for i, h in enumerate(header)}
    n = 0
    for r in body:
        r += ["", ""]
        v = ver.get(r[0])
        if not v:
            continue
        n += 1
        r[idx["ព័ត៌មានកែតម្រូវ"]] = corrected(v)
        if v.get("author_group"):
            r[idx["ក្រុមអ្នកនិពន្ធ"]] = v["author_group"]
        r[idx["ប្រភពផ្ទៀងផ្ទាត់"]] = " ; ".join(v.get("verification_urls") or [])
        r[idx["ស្ថានភាពផ្ទៀងផ្ទាត់"]] = v.get("status", "")
        r[idx["ស្ថានភាពអាន"]] = v.get("reading_status", "")
        note = r[idx["មូលហេតុ/កំណត់សម្គាល់"]]
        if v.get("correction"):
            note = (note + " | " if note else "") + "ការកែតម្រូវ៖ " + v["correction"]
        r[idx["មូលហេតុ/កំណត់សម្គាល់"]] = note
        r[idx["វិធីផ្ទៀងផ្ទាត់"]] = METHOD
        r[idx["សេចក្តីពិពណ៌នាបឋម (ពីអ្វីដែលបានឃើញ)"]] = v.get("relevance_note_km", "")
    with open(os.path.join(DATA, "u300-audit-v2.csv"), "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows([header] + body)

    arows = list(csv.reader(open(os.path.join(DATA, "a-candidates.csv"), encoding="utf-8-sig")))
    aheader = arows[0] + ["ព័ត៌មានកែតម្រូវ", "ប្រភពផ្ទៀងផ្ទាត់", "វិធីផ្ទៀងផ្ទាត់", "កំណត់សម្គាល់"]
    abody = []
    for r in arows[1:] + NEW_A:
        r = r + ["", "", "", ""]
        v = ver.get(r[0])
        if v:
            r[5] = v.get("status", "")
            r[6] = v.get("reading_status", "")
            r[7] = corrected(v)
            r[8] = " ; ".join(v.get("verification_urls") or [])
            r[9] = METHOD
            r[10] = v.get("correction", "")
        abody.append(r)
    with open(os.path.join(DATA, "a-candidates-v2.csv"), "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows([aheader] + abody)
    print(f"U rows updated: {n}; A rows: {len(abody)} ({sum(1 for r in abody if r[0] in ver)} verified records)")


if __name__ == "__main__":
    main()
