#!/usr/bin/env bash
set -euo pipefail

# Render installs OS packages declared in the repository-root apt.txt
# before running this build command. Do not run apt-get here: Render's
# build filesystem may be read-only for apt package lists.
python -m pip install -r requirements-kb.txt

python - <<'PY'
import fitz
import pytesseract

languages = pytesseract.get_languages(config="")
print("PyMuPDF:", fitz.VersionBind)
print("Tesseract languages:", ", ".join(sorted(languages)))
if "chi_sim" not in languages:
    raise SystemExit(
        "缺少 chi_sim 简体中文 OCR 语言包。请确认仓库根目录 apt.txt "
        "包含 tesseract-ocr 和 tesseract-ocr-chi-sim。"
    )
PY
