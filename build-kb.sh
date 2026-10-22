#!/usr/bin/env bash
set -euo pipefail

# Install the core service dependencies first.
python -m pip install -r requirements-kb.txt

# Optional portable OCR fallback. This avoids requiring apt-get/Tesseract binaries
# in Render's native Python build environment. Failure here must not block service deployment.
if python -m pip install 'rapidocr>=3.10.0,<4.0' 'onnxruntime>=1.20.0,<2.0'; then
    python - <<'PY'
try:
    import onnxruntime
    from rapidocr import RapidOCR
    import numpy as np
    from PIL import Image, ImageDraw

    # Exercise one inference call so missing ONNX Runtime/model errors are visible
    # during build rather than only after a user uploads a scanned PDF.
    engine = RapidOCR()
    test_image = Image.new("RGB", (320, 96), "white")
    ImageDraw.Draw(test_image).text((12, 28), "OCR TEST 123", fill="black")
    engine(np.asarray(test_image))
    print("RapidOCR runtime smoke test: inference call succeeded; ONNX Runtime", onnxruntime.__version__)
except Exception as exc:
    print("WARNING: RapidOCR runtime smoke test failed:", type(exc).__name__, str(exc)[:300])
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
