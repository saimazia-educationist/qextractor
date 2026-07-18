FROM python:3.12-slim

# tesseract-ocr is the actual OCR engine that pytesseract calls out to.
# libgl1/libglib2.0-0 are common runtime deps for PDF/image libs (PyMuPDF/Pillow).
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py question_extractor.py ./
COPY frontend ./frontend

# workdir/ (per-session files) and extraction_cache/ (OCR cache) should be
# mounted as a persistent volume in production - see deployment notes.
RUN mkdir -p workdir extraction_cache papers_library

EXPOSE 5000

# Single worker: OCR is CPU-heavy and the app already isolates users by
# session folder, so one worker with several threads is a safer default
# than several workers each loading OCR models. Raise --workers if your
# host has multiple cores and you've tested memory headroom.
# Long timeout: OCR on a multi-page scanned PDF can take a while.
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "1", "--threads", "4", "--timeout", "300", "app:app"]
