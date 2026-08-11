# DOCX Quiz Parser

Flask web version of `Quiz_formatter.py`. Parses WGU-style Word quiz documents
into text2qti format, which then feeds the Canvas Quiz Converter.

Handles multiple choice, true/false, essay, fill-in-the-blank, and matching
questions. Correct answers are detected from yellow highlighting in the .docx or
from answer-key lines ("Ans: D", "Answer - A", "Correct Answer: B").

## Files

| File | Purpose |
| --- | --- |
| `app.py` | Flask wrapper plus the embedded page HTML (PAGE constant) |
| `quiz_formatter.py` | The parser itself (unchanged logic; Tkinter GUI still works locally) |
| `requirements.txt` | Flask, python-docx, gunicorn |
| `.python-version` | Pins Python 3.12 on Render |

## Run locally

    pip install -r requirements.txt
    python app.py          # http://127.0.0.1:8012

The original desktop modes still work:

    python quiz_formatter.py              # Tkinter GUI
    python quiz_formatter.py file.docx    # command line

## Render settings

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app --timeout 120`
