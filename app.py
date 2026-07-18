"""
Question Extractor - Web backend.

Wraps the original desktop extraction logic (question_extractor.py) behind
a small HTTP API so it can run as a website instead of a PyQt5 desktop app.

Each browser session gets its own isolated working folder under WORKDIR, so
multiple users can use the site at once without stepping on each other's files.
"""
import os
import uuid
import shutil
import hashlib
import logging
from pathlib import Path

from flask import Flask, request, jsonify, send_file, session

from question_extractor import (
    extract_questions,
    generate_test,
    renumber_questions,
    get_chapter_list,
    InvalidPaperError,
)

logging.basicConfig(level=logging.INFO)

APP_ROOT = Path(__file__).parent
WORKDIR = APP_ROOT / "workdir"
WORKDIR.mkdir(exist_ok=True)

# Extraction (OCR + PDF parsing) is the slow part of this app. Since many
# users upload the exact same well-known past papers, we cache extraction
# results by file content hash - so a paper only ever needs to be OCR'd
# once, ever, regardless of who uploads it or how many times.
CACHE_DIR = APP_ROOT / "extraction_cache"
CACHE_DIR.mkdir(exist_ok=True)

# Drop PDF past papers directly into this folder on the server (named the
# normal CIE way, e.g. 9608_s15_qp_11.pdf) and they'll show up as a
# pick-from-list option, so visitors don't need their own copies on disk.
PAPERS_LIBRARY_DIR = APP_ROOT / "papers_library"
PAPERS_LIBRARY_DIR.mkdir(exist_ok=True)

MAX_UPLOAD_MB = 100

app = Flask(__name__, static_folder=str(APP_ROOT.parent / "frontend"), static_url_path="")

_secret = os.environ.get("FLASK_SECRET_KEY")
if not _secret:
    if os.environ.get("FLASK_ENV") == "production" or os.environ.get("RENDER"):
        raise RuntimeError(
            "FLASK_SECRET_KEY environment variable must be set in production."
        )
    _secret = "dev-secret-change-me"
    logging.warning("Using insecure default FLASK_SECRET_KEY - set FLASK_SECRET_KEY env var in production.")
app.secret_key = _secret

app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

PAPER_NAMES = ["Paper 1", "Paper 2", "Paper 3", "Paper 4"]


def get_session_dir() -> Path:
    """Each visitor gets a private folder, identified by a cookie session id."""
    if "sid" not in session:
        session["sid"] = uuid.uuid4().hex
    sdir = WORKDIR / session["sid"]
    sdir.mkdir(exist_ok=True)
    return sdir


def file_sha256(path: Path) -> str:
    """Content hash used as the cache key - same PDF content = same hash,
    regardless of what the file happens to be named."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def merge_docx_folder(source_dir: Path, dest_dir: Path):
    """
    Copies every chapter .docx from source_dir into dest_dir. If dest_dir
    already has a file with that chapter's name (because an earlier PDF in
    this same batch already contributed questions to it), the new content
    is appended rather than overwriting - so combining multiple papers'
    results (cached or freshly extracted) still merges correctly by topic.
    """
    from docx import Document
    from docxcompose.composer import Composer

    dest_dir.mkdir(parents=True, exist_ok=True)
    for src_file in source_dir.glob("*.docx"):
        dest_file = dest_dir / src_file.name
        if dest_file.exists():
            master = Document(str(dest_file))
            composer = Composer(master)
            composer.append(Document(str(src_file)))
            composer.save(str(dest_file))
        else:
            shutil.copy(src_file, dest_file)


def get_chapters_for_pdf(pdf_path: Path) -> Path:
    """
    Returns a folder containing the per-topic .docx files extracted from
    this single PDF - from cache if this exact file has been extracted
    before (by anyone), otherwise runs the real OCR/extraction once and
    saves the result to the cache for next time.
    """
    file_hash = file_sha256(pdf_path)
    cache_entry = CACHE_DIR / file_hash
    marker = cache_entry / "_complete"

    if marker.exists():
        logging.info(f"Cache hit for {pdf_path.name} ({file_hash[:12]}...)")
        return cache_entry

    logging.info(f"Cache miss for {pdf_path.name} ({file_hash[:12]}...) - running extraction")
    cache_entry.mkdir(parents=True, exist_ok=True)
    try:
        extract_questions([str(pdf_path)], "new", str(cache_entry))
        marker.touch()  # only mark complete on success - a failed/partial run gets retried next time
    except Exception:
        shutil.rmtree(cache_entry, ignore_errors=True)
        raise

    return cache_entry


def parse_paper_meta(filename: str) -> dict:
    """
    Reads the paper number/session out of a CIE-style filename, e.g.
    '9608_s15_qp_11.pdf' -> subject 9608, session 'May/June 2015', Paper 1.
    Purely for display in the library picker - doesn't affect extraction.
    """
    basename = os.path.splitext(filename)[0]
    parts = basename.split("_")
    subject = parts[0] if parts else "?"
    session_code = parts[1] if len(parts) > 1 else ""
    doc_type = parts[2].lower() if len(parts) > 2 else ""  # 'qp' = question paper, 'ms' = mark scheme
    paper_code = parts[-1] if parts else ""

    label = "Unknown paper"
    for codes, name in (
        (("11", "12", "13"), "Paper 1"),
        (("21", "22", "23"), "Paper 2"),
        (("31", "32", "33"), "Paper 3"),
        (("41", "42", "43"), "Paper 4"),
    ):
        if paper_code in codes:
            label = name
            break

    season_map = {"s": "May/June", "w": "Oct/Nov", "m": "Feb/March"}
    season = season_map.get(session_code[:1].lower(), "") if session_code else ""
    year = f"20{session_code[1:]}" if len(session_code) > 1 and session_code[1:].isdigit() else ""
    session_display = f"{season} {year}".strip() or "Unknown session"

    return {"subject": subject, "paper_label": label, "session": session_display, "doc_type": doc_type}


@app.route("/api/library")
def api_library():
    """
    Lists question papers (qp) the site owner has placed in papers_library/,
    for the pick-from-list option. Mark schemes (ms) are excluded for now -
    that's a separate feature to be added later.
    """
    papers = []
    for pdf in sorted(PAPERS_LIBRARY_DIR.glob("*.pdf")):
        meta = parse_paper_meta(pdf.name)
        if meta["doc_type"] != "qp":
            continue
        papers.append({"filename": pdf.name, **meta})
    return jsonify({"papers": papers})


@app.route("/api/chapters")
def api_chapters():
    """Returns available chapter names for every paper, for the selection UI."""
    return jsonify({paper: get_chapter_list(paper) for paper in PAPER_NAMES})


@app.route("/api/extract", methods=["POST"])
def api_extract():
    """
    Accepts one or more PDF question papers - either uploaded directly, or
    picked from the server's papers_library/ - and writes per-topic .docx
    files into this session's folder. Each PDF is checked against the
    extraction cache first (by content hash) - only genuinely new papers
    trigger real OCR/extraction.
    """
    files = request.files.getlist("pdfs")
    library_selected = request.form.getlist("library_files")

    sdir = get_session_dir()
    upload_dir = sdir / "uploads"
    chapter_dir = sdir / "chapters"
    upload_dir.mkdir(exist_ok=True)

    mode = "new" if request.form.get("mode", "new") == "new" else "append"
    if mode == "new":
        shutil.rmtree(chapter_dir, ignore_errors=True)
    chapter_dir.mkdir(exist_ok=True)

    saved_paths = []
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            return jsonify({"error": f"'{f.filename}' is not a PDF."}), 400
        dest = upload_dir / f.filename
        f.save(dest)
        saved_paths.append(dest)

    for name in library_selected:
        safe_name = os.path.basename(name)
        candidate = PAPERS_LIBRARY_DIR / safe_name
        if not candidate.exists():
            return jsonify({"error": f"'{safe_name}' was not found in the paper library."}), 400
        saved_paths.append(candidate)

    if not saved_paths:
        return jsonify({"error": "No PDF files selected."}), 400

    cache_hits = 0
    try:
        for pdf_path in saved_paths:
            per_file_dir = get_chapters_for_pdf(pdf_path)
            if (per_file_dir / "_complete").exists():
                cache_hits += 1
            merge_docx_folder(per_file_dir, chapter_dir)
    except InvalidPaperError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logging.exception("Extraction failed")
        return jsonify({"error": f"Extraction failed: {e}"}), 500

    produced = sorted(p.name for p in chapter_dir.glob("*.docx"))
    return jsonify({
        "chapters_available": produced,
        "cache_hits": cache_hits,
        "total_files": len(saved_paths),
    })


@app.route("/api/download-extracted")
def api_download_extracted():
    """
    Zips up every per-topic .docx file produced by the last extraction in
    this session, so the user can keep them - mirrors the desktop app's
    'pick an output folder' behaviour, just as a downloadable zip instead.
    """
    sdir = get_session_dir()
    chapter_dir = sdir / "chapters"
    docx_files = list(chapter_dir.glob("*.docx")) if chapter_dir.exists() else []
    if not docx_files:
        return jsonify({"error": "Nothing extracted yet."}), 404

    zip_base = sdir / "extracted_questions"
    zip_path = shutil.make_archive(str(zip_base), "zip", root_dir=str(chapter_dir))
    return send_file(zip_path, as_attachment=True, download_name="extracted_questions.zip")


@app.route("/api/create-test", methods=["POST"])
def api_create_test():
    """
    Builds a combined test .docx. Two source modes, matching the desktop app:

    - source=papers: uploaded PDF question papers are freshly extracted into
      a scratch folder, then filtered down to the selected chapters.
    - source=docx: previously-extracted .docx chapter files are uploaded
      directly (what a user would have gotten from /api/download-extracted
      earlier), then filtered down to the selected chapters.
    """
    source = request.form.get("source", "papers")
    selected = request.form.getlist("chapters")  # may be empty -> use everything uploaded
    sdir = get_session_dir()

    scratch = sdir / f"test_build_{uuid.uuid4().hex}"
    scratch.mkdir(parents=True, exist_ok=True)

    try:
        if source == "papers":
            pdf_files = request.files.getlist("pdfs")
            library_selected = request.form.getlist("library_files")

            upload_dir = scratch / "uploads"
            chapter_dir = scratch / "chapters"
            upload_dir.mkdir(exist_ok=True)
            chapter_dir.mkdir(exist_ok=True)

            saved_paths = []
            for f in pdf_files:
                if not f.filename.lower().endswith(".pdf"):
                    return jsonify({"error": f"'{f.filename}' is not a PDF."}), 400
                dest = upload_dir / f.filename
                f.save(dest)
                saved_paths.append(dest)

            for name in library_selected:
                safe_name = os.path.basename(name)
                candidate = PAPERS_LIBRARY_DIR / safe_name
                if not candidate.exists():
                    return jsonify({"error": f"'{safe_name}' was not found in the paper library."}), 400
                saved_paths.append(candidate)

            if not saved_paths:
                return jsonify({"error": "No question paper PDFs selected."}), 400

            for pdf_path in saved_paths:
                per_file_dir = get_chapters_for_pdf(pdf_path)
                merge_docx_folder(per_file_dir, chapter_dir)

            all_produced = list(chapter_dir.glob("*.docx"))

        elif source == "docx":
            docx_files = request.files.getlist("docx_files")
            if not docx_files:
                return jsonify({"error": "No extracted .docx files uploaded."}), 400

            chapter_dir = scratch / "chapters"
            chapter_dir.mkdir(exist_ok=True)
            for f in docx_files:
                if not f.filename.lower().endswith(".docx"):
                    continue
                f.save(chapter_dir / os.path.basename(f.filename))

            all_produced = list(chapter_dir.glob("*.docx"))

        else:
            return jsonify({"error": "Unknown source type."}), 400

        if not all_produced:
            return jsonify({"error": "No chapter files were available to build a test from."}), 400

        if selected:
            selected_names = {os.path.basename(n) for n in selected}
            chapter_paths = [str(p) for p in all_produced if p.name in selected_names]
            if not chapter_paths:
                return jsonify({"error": "None of the selected chapters matched the available files. Using all instead."}), 400
        else:
            chapter_paths = [str(p) for p in all_produced]

        output_path = sdir / "test.docx"
        generate_test(chapter_paths, str(output_path))
        renumber_questions(str(output_path))

    except InvalidPaperError as e:
        shutil.rmtree(scratch, ignore_errors=True)
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logging.exception("Test generation failed")
        return jsonify({"error": f"Test generation failed: {e}"}), 500
    finally:
        shutil.rmtree(scratch, ignore_errors=True)

    return jsonify({"ready": True, "download_url": "/api/download"})


@app.route("/api/download")
def api_download():
    sdir = get_session_dir()
    output_path = sdir / "test.docx"
    if not output_path.exists():
        return jsonify({"error": "No test has been generated yet."}), 404
    return send_file(output_path, as_attachment=True, download_name="test.docx")


@app.route("/api/reset", methods=["POST"])
def api_reset():
    """Clears this session's working files so the user can start fresh."""
    sdir = get_session_dir()
    shutil.rmtree(sdir, ignore_errors=True)
    sdir.mkdir(exist_ok=True)
    return jsonify({"ok": True})


@app.route("/")
def index():
    return app.send_static_file("index.html")


if __name__ == "__main__":
    debug_mode = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug_mode, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
