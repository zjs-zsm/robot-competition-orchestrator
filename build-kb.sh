#!/usr/bin/env bash
set -euo pipefail

if ! command -v apt-get >/dev/null 2>&1; then
  echo "错误：当前构建环境没有 apt-get，无法自动安装 Tesseract。" >&2
  echo "请在支持 Debian/Ubuntu 包管理器的 Render Python 环境中运行此构建脚本。" >&2
  exit 1
fi

apt-get update
DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
  tesseract-ocr tesseract-ocr-chi-sim
rm -rf /var/lib/apt/lists/*

python -m pip install -r requirements-kb.txt
python - <<'PY'
import fitz
import pytesseract
languages = pytesseract.get_languages(config="")
print("PyMuPDF:", fitz.VersionBind)
print("Tesseract languages:", ", ".join(sorted(languages)))
if "chi_sim" not in languages:
    raise SystemExit("缺少 chi_sim 简体中文 OCR 语言包，PDF OCR 无法工作。")
PY
