# សៀវភៅ «ចរិត និងមរតកវប្បធម៌នៃការដឹកនាំខ្មែរ» — ដំណាក់កាលទី១

គម្រោងសៀវភៅនេះគឺជាដំណាក់កាលរៀបចំផែនការ (Phase 1) ដែលធ្វើតាម `../prompts/Claude-Prompt-Complete-Khmer-v2.txt`។ **មិនទាន់មានខ្លឹមសារសៀវភៅទេ** — ការសរសេរនឹងចាប់ផ្តើមតែក្រោយពេលអ្នកនិពន្ធយល់ព្រមលើគម្រោង។

## ឯកសារ

| ផ្លូវ | ខ្លឹមសារ |
|---|---|
| `phase1-plan.docx` | របាយការណ៍ដំណាក់កាលទី១ (Word; Khmer OS Siemreap / Khmer OS Muol Light) |
| `phase1-plan-preview.pdf` | ការមើលជាមុនជា PDF (បង្កើតដោយពុម្ព Siemreap/Moul ពី Google Fonts) |
| `phase1/*.md` | អត្ថបទប្រភពនៃរបាយការណ៍ (លំដាប់ ១–១២ តាមផ្នែក ១៣ របស់ Prompt) |
| `data/source-sections.csv` | ផ្នែកទាំង ៤១ នៃឯកសារគោល ជាមួយចំនួនពាក្យ/តួអក្សរ |
| `data/leader-roster.csv` | បញ្ជីអ្នកដឹកនាំ L001–L045 |
| `data/case-studies.csv` | ករណីសិក្សាស៊ីជម្រៅ C01–C14 |
| `data/u300-audit.csv` | តារាងសវនកម្មបឋម U001–U300 (មិនទាន់ផ្ទៀងផ្ទាត់) |
| `data/a-candidates.csv` | បេក្ខភាពប្រភពបន្ថែម A001–A023 (មិនទាន់ផ្ទៀងផ្ទាត់) |
| `tools/count_source_words.py` | រាប់ពាក្យខ្មែរតាមផ្នែក (khmer-nltk) |
| `tools/make_data.py` | បង្កើតតារាង CSV និងតារាងដែលបង្កើតដោយកម្មវិធីក្នុង `phase1/` |
| `build.py` | បង្កើត `phase1-plan.docx` ពី `phase1/*.md` (ប្រើកូដរបស់ `../dissertation/build.py`) |

ឯកសារគោលដែលអ្នកនិពន្ធបានផ្តល់ (.docx) **មិនត្រូវបានដាក់ក្នុងឃ្លាំងនេះទេ** ដោយសារបញ្ហាសិទ្ធិអ្នកនិពន្ធ (មើលបញ្ជីហានិភ័យ R11)។

## ការបង្កើតឡើងវិញ

```bash
pip install khmer-nltk python-docx      # pandoc ត្រូវតែមាន
pandoc source.docx -t markdown -o source.md
python3 -I tools/count_source_words.py source.md > tools/source_sections.json
python3 tools/make_data.py
python3 build.py                         # ឬ --font Siemreap --heading-font Moul
```
