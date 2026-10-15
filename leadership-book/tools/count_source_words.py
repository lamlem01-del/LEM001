"""Count Khmer words and characters per heading-1 section of a docx/markdown source.

Usage:
    pandoc source.docx -t markdown -o source.md
    python3 -I count_source_words.py source.md > sections.json

Word segmentation uses khmer-nltk (CRF model; `pip install khmer-nltk`). Zero-width
spaces and ordinary spaces are removed before segmentation, so results do not depend
on how the typist spaced the text. Tokens without a Khmer letter (punctuation, digits,
Latin) are not counted as words. Characters = Khmer code points U+1780–U+17DD.
"""
import json
import logging
import re
import sys

logging.disable(logging.CRITICAL)
from khmernltk import word_tokenize  # noqa: E402

KHMER_RUN = re.compile(r"[ក-៝​ ]+")
KHMER_LETTER = re.compile(r"[ក-៓]")


def count(text):
    words = 0
    for chunk in KHMER_RUN.findall(text):
        chunk = chunk.replace(" ", "").replace("​", "")
        if chunk:
            words += sum(1 for t in word_tokenize(chunk, return_tokens=True) if KHMER_LETTER.search(t))
    chars = len(re.findall(r"[ក-៝]", text))
    return words, chars


def main(path):
    md = open(path, encoding="utf-8").read()
    md = re.sub(r"!\[[^\]]*\]\([^)]*\)(\{[^}]*\})?", "", md)  # images
    md = re.sub(r"\]\(https?://[^)]*\)", "", md)  # link targets
    parts = re.split(r"^# (.*)$", md, flags=re.M)
    out = []
    for i in range(1, len(parts), 2):
        words, chars = count(parts[i + 1])
        out.append({"i": len(out) + 1, "heading": parts[i].strip(), "words": words, "chars": chars})
    json.dump(out, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(sys.argv[1])
