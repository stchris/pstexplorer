# Named Entity Recognition over an exported PST SQLite database.
#
# Usage:
#   cargo run -- export my.pst          # produces my.db
#   uv run --python 3.12 --with spacy --with "pydantic<2" --with compressed-rtf --with striprtf ner.py my.db
#
# First run only — download the spaCy model:
#   uvx --python 3.12 --with pip --with spacy --with "pydantic<2" python -m spacy download en_core_web_lg

import re
import sqlite3
import sys
from collections import Counter

import compressed_rtf
import spacy
from striprtf.striprtf import rtf_to_text


def html_strip(html: str) -> str:
    return re.sub(r"<[^>]+>", " ", html)


def rtf_blob_to_text(blob: bytes) -> str:
    try:
        rtf_bytes = compressed_rtf.decompress(blob)
        return rtf_to_text(rtf_bytes.decode("utf-8", errors="replace"))
    except Exception:
        return ""


def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <database.db>", file=sys.stderr)
        sys.exit(1)

    db = sqlite3.connect(sys.argv[1])
    nlp = spacy.load("en_core_web_lg")

    rows = db.execute(
        """
        SELECT sender, to_recipients, subject, body_text, body_html, body_rtf
        FROM messages
        WHERE message_class LIKE 'IPM.NOTE%'
           OR message_class = 'IPM'
           OR message_class = ''
        """
    ).fetchall()

    print(f"Processing {len(rows)} messages...", file=sys.stderr)

    entities: Counter = Counter()

    for sender, to, subject, body_text, body_html, body_rtf in rows:
        if body_text:
            body = body_text
        elif body_html:
            body = html_strip(body_html)
        elif body_rtf:
            body = rtf_blob_to_text(bytes(body_rtf))
        else:
            body = ""

        text = " ".join(filter(None, [sender, to, subject, body[:8000]]))
        if not text.strip():
            continue

        doc = nlp(text)
        for ent in doc.ents:
            cleaned = ent.text.strip()
            if cleaned:
                entities[(ent.label_, cleaned)] += 1

    print(f"\n{'Count':>6}  {'Label':<12}  Entity")
    print("-" * 60)
    for (label, text), count in entities.most_common(100):
        print(f"{count:6d}  {label:<12}  {text}")


if __name__ == "__main__":
    main()
