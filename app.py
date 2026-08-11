"""
app.py — Flask wrapper around quiz_formatter.py for Render hosting.

Local:  python app.py            → http://127.0.0.1:8012
Render: gunicorn app:app --timeout 120
"""

import io
import os
import tempfile
import zipfile
from pathlib import Path

from flask import Flask, render_template, request, send_file

from quiz_formatter import parse_docx_to_text2qti

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB per request


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
    return render_template("index.html", results=results)


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
