#!/usr/bin/env bash
set -euo pipefail

# Install Python dependencies. Tesseract is an optional system dependency on
# Render native runtimes; do not fail the entire service build if it is absent.
python -m pip install -r requirements-kb.txt

python - <<'PY'
import fitz
print("PyMuPDF:", fitz.VersionBind)
try:
    import pytesseract
    languages = sorted(pytesseract.get_languages(config=""))
    print("Tesseract languages:", ", ".join(languages))
    if "chi_sim" not in languages:
        print("WARNING: Simplified Chinese OCR is unavailable; searchable embedded PDF text will still be extracted.")
except Exception as exc:
    print("WARNING: Tesseract OCR is unavailable:", type(exc).__name__)
    print("PDFs with embedded text can still be indexed; image-only/scanned PDFs may need OCR support.")
PY
