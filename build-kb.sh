#!/usr/bin/env bash
set -euo pipefail

# Install the core service dependencies first.
python -m pip install -r requirements-kb.txt

# Optional portable OCR fallback. This avoids requiring apt-get/Tesseract binaries
# in Render's native Python build environment. Failure here must not block service deployment.
if python -m pip install 'rapidocr>=3.10.0,<4.0'; then
    python - <<'PY'
try:
    from rapidocr import RapidOCR
    print("RapidOCR fallback: import successful")
except Exception as exc:
    print("WARNING: RapidOCR fallback could not be imported:", type(exc).__name__)
PY
else
    echo "WARNING: RapidOCR fallback installation failed; embedded PDF text extraction remains available."
fi

python - <<'PY'
import fitz
print("PyMuPDF:", fitz.VersionBind)
try:
    import pytesseract
    languages = sorted(pytesseract.get_languages(config=""))
    print("Tesseract languages:", ", ".join(languages))
    if "chi_sim" not in languages:
        print("WARNING: Simplified Chinese Tesseract OCR is unavailable; RapidOCR will be tried for scanned PDFs.")
except Exception as exc:
    print("WARNING: Tesseract OCR is unavailable:", type(exc).__name__)
    print("PDFs with embedded text can still be indexed; RapidOCR will be tried for image-only/scanned PDFs.")
PY
