"""
app.py — Flask wrapper around quiz_formatter.py for Render hosting.

Single-file version: the page HTML lives in the PAGE constant below, so there
is no templates/ folder to misplace.

Local:  python app.py            → http://127.0.0.1:8012
Render: gunicorn app:app --timeout 120
"""

import io
import os
import tempfile
import zipfile
from pathlib import Path

from flask import Flask, render_template_string, request, send_file

from quiz_formatter import parse_docx_to_text2qti

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB per request

PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DOCX Quiz Parser</title>
<style>
  :root {
    --ink: #1c2733;
    --paper: #f7f5f0;
    --card: #ffffff;
    --rule: #d9d4c8;
    --accent: #1f6f5c;
    --accent-soft: #e6f0ec;
    --warn: #8a4b1f;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    padding: 2.5rem 1.25rem 4rem;
    background: var(--paper);
    color: var(--ink);
    font: 16px/1.55 "Segoe UI", -apple-system, BlinkMacSystemFont, Helvetica, Arial, sans-serif;
  }
  .wrap { max-width: 900px; margin: 0 auto; }
  h1 {
    font-size: 1.6rem;
    letter-spacing: -0.01em;
    margin: 0 0 .25rem;
  }
  .sub { color: #5d6b78; margin: 0 0 1.75rem; font-size: .95rem; }
  .card {
    background: var(--card);
    border: 1px solid var(--rule);
    border-radius: 10px;
    padding: 1.25rem 1.25rem 1.4rem;
    margin-bottom: 1.5rem;
  }
  label { display: block; font-weight: 600; font-size: .9rem; margin-bottom: .5rem; }
  input[type=file] {
    width: 100%;
    padding: .7rem;
    border: 1px dashed var(--rule);
    border-radius: 8px;
    background: #fbfaf7;
  }
  .actions { display: flex; gap: .6rem; flex-wrap: wrap; margin-top: 1rem; }
  button {
    font: inherit;
    font-weight: 600;
    padding: .6rem 1.1rem;
    border-radius: 8px;
    border: 1px solid var(--accent);
    background: var(--accent);
    color: #fff;
    cursor: pointer;
  }
  button.secondary { background: #fff; color: var(--accent); }
  button:hover { filter: brightness(1.06); }
  button:focus-visible { outline: 3px solid #94c4b6; outline-offset: 2px; }
  h2 { font-size: 1.05rem; margin: 0 0 .6rem; }
  .file-name {
    display: inline-block;
    background: var(--accent-soft);
    color: var(--accent);
    border-radius: 5px;
    padding: .15rem .5rem;
    font-size: .85rem;
    font-weight: 600;
  }
  pre {
    background: #fbfaf7;
    border: 1px solid var(--rule);
    border-radius: 8px;
    padding: .9rem;
    overflow: auto;
    max-height: 26rem;
    font: 13px/1.5 "SFMono-Regular", Consolas, "Liberation Mono", monospace;
    white-space: pre-wrap;
    word-break: break-word;
  }
  .report { background: var(--accent-soft); border-color: #c2ddd4; }
  .error { color: var(--warn); font-weight: 600; }
  .hint { font-size: .85rem; color: #5d6b78; margin-top: .6rem; }
</style>
</head>
<body>
<div class="wrap">
  <h1>DOCX Quiz Parser</h1>
  <p class="sub">Turns WGU-style Word quiz files into text2qti text, ready for the Canvas Quiz Converter.</p>

  <form method="post" enctype="multipart/form-data" class="card">
    <label for="docx">Word quiz files (.docx)</label>
    <input id="docx" type="file" name="docx" accept=".docx" multiple required>
    <div class="actions">
      <button type="submit">Parse files</button>
      <button type="submit" class="secondary" formaction="/download">Parse and download</button>
    </div>
    <p class="hint">Correct answers are read from yellow highlighting or answer-key lines such as "Ans: D" or "Correct Answer: B". Several files can be parsed at once; downloading more than one returns a .zip.</p>
  </form>

  {% for r in results %}
  <div class="card">
    <h2>{% if r.name %}<span class="file-name">{{ r.name }}</span>{% else %}Nothing to parse{% endif %}</h2>
    {% if r.error %}
      <p class="error">{{ r.error }}</p>
    {% else %}
      <pre class="report">{{ r.report }}</pre>
      <pre>{{ r.text }}</pre>
    {% endif %}
  </div>
  {% endfor %}
</div>
</body>
</html>
"""


def _parse_upload(storage):
    """Save an uploaded .docx to a temp file, parse it, clean up."""
    suffix = Path(storage.filename).suffix or ".docx"
    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    try:
        storage.save(tmp.name)
        tmp.close()
        text, report = parse_docx_to_text2qti(tmp.name)
        return {
            "name": Path(storage.filename).stem,
            "text": text,
            "report": report,
            "error": None,
        }
    except Exception as exc:  # noqa: BLE001 — surfaced to the user in the results panel
        return {
            "name": Path(storage.filename).stem,
            "text": "",
            "report": "",
            "error": f"{type(exc).__name__}: {exc}",
        }
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    if request.method == "POST":
        files = [f for f in request.files.getlist("docx") if f and f.filename]
        if not files:
            results.append(
                {"name": "", "text": "", "report": "", "error": "Choose at least one .docx file."}
            )
        for f in files:
            results.append(_parse_upload(f))
    return render_template_string(PAGE, results=results)


@app.route("/download", methods=["POST"])
def download():
    """Re-parse the posted files and return .txt (one file) or .zip (several)."""
    files = [f for f in request.files.getlist("docx") if f and f.filename]
    if not files:
        return "No files posted.", 400

    parsed = [_parse_upload(f) for f in files]
    good = [p for p in parsed if not p["error"]]
    if not good:
        return "Nothing parsed successfully.", 400

    if len(good) == 1:
        buf = io.BytesIO(good[0]["text"].encode("utf-8"))
        return send_file(
            buf,
            mimetype="text/plain",
            as_attachment=True,
            download_name=f"{good[0]['name']}_text2qti.txt",
        )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in good:
            z.writestr(f"{p['name']}_text2qti.txt", p["text"])
    buf.seek(0)
    return send_file(
        buf, mimetype="application/zip", as_attachment=True, download_name="text2qti_files.zip"
    )


@app.route("/healthz")
def healthz():
    return "ok", 200


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 8012)), debug=True)
