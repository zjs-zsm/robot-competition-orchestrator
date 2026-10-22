from __future__ import annotations

import gc
import io
import logging
import os
import re
import threading

logger = logging.getLogger("robot_competition_knowledge_base.pdf")

# One small OCR engine per process; serialize scanned-PDF work to avoid duplicate models
# exhausting the memory limit when two reindex/upload requests arrive together.
_RAPIDOCR_ENGINE = None
_RAPIDOCR_LOCK = threading.Lock()


def pdf_text_is_suspicious(text: str) -> bool:
    """Detect common PDF font-map corruption before indexing unreadable text."""
    if not text or not text.strip():
        return True

    compact = re.sub(r"\s+", "", text)
    if not compact:
        return True

    bad_chars = sum(1 for ch in compact if ch == "\ufffd" or "\ue000" <= ch <= "\uf8ff")
    if bad_chars / len(compact) > 0.005:
        return True

    controls = sum(
        1 for ch in text
        if ord(ch) < 32 and ch not in "\n\r\t\f"
    )
    if controls / max(1, len(text)) > 0.005:
        return True

    cjk_count = len(re.findall(r"[\u3400-\u9fff]", text))
    long_upper_tokens = re.findall(r"(?<![A-Za-z0-9])[A-Z][A-Z0-9]{5,}(?![A-Za-z0-9])", text)
    if cjk_count >= 20 and len(long_upper_tokens) >= 4:
        token_chars = sum(len(token) for token in long_upper_tokens)
        repetition_ratio = 1 - (len(set(long_upper_tokens)) / len(long_upper_tokens))
        # Broken ToUnicode maps often turn Chinese PDF text into repeated uppercase codes.
        if repetition_ratio >= 0.30 and token_chars >= 24:
            return True
        if token_chars >= max(60, cjk_count * 0.55):
            return True

    return False


def _score_text(text: str) -> float:
    compact = re.sub(r"\s+", "", text or "")
    cjk_count = len(re.findall(r"[\u3400-\u9fff]", text or ""))
    score = min(len(compact), 5000) / 100 + min(cjk_count, 1000) / 20
    if pdf_text_is_suspicious(text):
        score -= 1000
    return score


def _ocr_pdf(raw: bytes) -> tuple[str, str]:
    """OCR pages when Tesseract and its Simplified Chinese language pack are installed."""
    try:
        import fitz
        import pytesseract
        from PIL import Image
    except ImportError:
        return "", "OCR依赖未安装"

    try:
        languages = set(pytesseract.get_languages(config=""))
    except Exception as exc:
        logger.warning("Tesseract is not available: %s", type(exc).__name__)
        return "", "OCR未启用"

    if "chi_sim" not in languages:
        return "", "缺少简体中文OCR语言包"

    try:
        configured_limit = int(os.getenv("MAX_PDF_OCR_PAGES", "40"))
        page_limit = max(1, min(configured_limit, 100))
        page_texts: list[str] = []
        with fitz.open(stream=raw, filetype="pdf") as pdf:
            page_count = len(pdf)
            for page_index in range(min(page_count, page_limit)):
                page = pdf.load_page(page_index)
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                with Image.open(io.BytesIO(pixmap.tobytes("png"))) as image:
                    page_text = pytesseract.image_to_string(
                        image,
                        lang="chi_sim+eng",
                        config="--psm 3",
                        timeout=20,
                    ).strip()
                if page_text:
                    page_texts.append(f"第{page_index + 1}页：\n{page_text}")

        result = "\n\n".join(page_texts).strip()
        if page_count > page_limit:
            logger.warning(
                "PDF has %s pages; OCR processed only the first %s pages. "
                "Set MAX_PDF_OCR_PAGES to increase the limit.",
                page_count,
                page_limit,
            )
        return result, "OCR识别"
    except Exception as exc:
        logger.warning("PDF OCR failed: %s", type(exc).__name__, exc_info=True)
        return "", "OCR失败"


def _get_rapidocr_engine():
    """Create one memory-conscious OCR engine and reuse it across PDF requests."""
    global _RAPIDOCR_ENGINE
    if _RAPIDOCR_ENGINE is None:
        from rapidocr import RapidOCR
        from rapidocr.utils.typings import ModelType

        _RAPIDOCR_ENGINE = RapidOCR(
            params={
                # Tiny PP-OCRv6 models plus no angle classifier reduce peak RAM on
                # Render's 512 MiB instance while retaining Chinese/English OCR.
                "Global.use_cls": False,
                "Global.max_side_len": 1600,
                "Det.model_type": ModelType.TINY,
                "Rec.model_type": ModelType.TINY,
                "EngineConfig.onnxruntime.intra_op_num_threads": 1,
                "EngineConfig.onnxruntime.inter_op_num_threads": 1,
            }
        )
    return _RAPIDOCR_ENGINE


def _rapidocr_pdf(raw: bytes) -> tuple[str, str]:
    """Memory-bounded Chinese/English OCR fallback without a system Tesseract binary."""
    try:
        import fitz
        import numpy as np
        from PIL import Image
    except Exception as exc:
        logger.warning("RapidOCR dependencies are unavailable: %s", type(exc).__name__)
        return "", "OCR未启用"

    try:
        configured_limit = int(os.getenv("MAX_PDF_OCR_PAGES", "40"))
        page_limit = max(1, min(configured_limit, 100))
        page_texts: list[str] = []

        # The lock covers model creation and the full document. Without it, concurrent
        # requests can each allocate their own ONNX sessions and kill the service.
        with _RAPIDOCR_LOCK:
            engine = _get_rapidocr_engine()
            with fitz.open(stream=raw, filetype="pdf") as pdf:
                page_count = len(pdf)
                for page_index in range(min(page_count, page_limit)):
                    page = pdf.load_page(page_index)
                    # 1.5x is sufficient for the detector's bounded input and uses less RAM.
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                    with Image.open(io.BytesIO(pixmap.tobytes("png"))) as image:
                        image_array = np.asarray(image.convert("RGB"))
                    result = engine(image_array)
                    texts = getattr(result, "txts", None)
                    if texts is None and isinstance(result, dict):
                        texts = result.get("txts")
                    # Compatibility with older RapidOCR versions returning
                    # (boxes, text, scores).
                    if texts is None and isinstance(result, (tuple, list)) and result:
                        first = result[0]
                        if isinstance(first, list):
                            texts = [
                                item[1] for item in first
                                if isinstance(item, (tuple, list)) and len(item) >= 2
                            ]
                    if texts:
                        page_text = "\n".join(
                            str(item).strip() for item in texts if str(item).strip()
                        )
                        if page_text:
                            page_texts.append(f"第{page_index + 1}页：\n{page_text}")

                    # Release page buffers before rendering the next page.
                    del result, image_array, pixmap
                    gc.collect()

                if page_count > page_limit:
                    logger.warning(
                        "PDF has %s pages; RapidOCR processed only the first %s pages.",
                        page_count,
                        page_limit,
                    )

        result_text = "\n\n".join(page_texts).strip()
        return result_text, "RapidOCR识别" if result_text else "OCR未识别到文字"
    except Exception as exc:
        logger.warning("RapidOCR failed: %s", type(exc).__name__, exc_info=True)
        return "", "OCR失败"

def extract_pdf_text(raw: bytes) -> tuple[str, str]:
    """
    Prefer reliable embedded text; compare pypdf and PyMuPDF extraction.
    Suspicious or image-only PDFs are sent through Chinese/English OCR when available.
    """
    candidates: list[tuple[str, str]] = []

    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(raw), strict=False)
        text = "\n".join(page.extract_text() or "" for page in reader.pages).strip()
        if text:
            candidates.append(("pypdf", text))
    except Exception as exc:
        logger.info("pypdf extraction failed: %s", type(exc).__name__)

    try:
        import fitz
        with fitz.open(stream=raw, filetype="pdf") as pdf:
            text = "\n".join(page.get_text("text", sort=True) or "" for page in pdf).strip()
        if text:
            candidates.append(("pymupdf", text))
    except Exception as exc:
        logger.info("PyMuPDF extraction failed: %s", type(exc).__name__)

    reliable = [
        (source, text)
        for source, text in candidates
        if text.strip() and not pdf_text_is_suspicious(text)
    ]
    best_reliable = max(reliable, key=lambda item: _score_text(item[1]), default=None)

    # Long, readable embedded text needs no OCR. Short text may be only a title or header.
    if best_reliable and len(re.sub(r"\s+", "", best_reliable[1])) >= 80:
        logger.info("PDF text extracted using %s.", best_reliable[0])
        return best_reliable[1], "已入库"

    ocr_text, ocr_reason = _ocr_pdf(raw)
    if not ocr_text.strip():
        # Tesseract may be absent on native Render runtimes; try the portable OCR engine.
        rapid_text, rapid_reason = _rapidocr_pdf(raw)
        if rapid_text.strip():
            ocr_text, ocr_reason = rapid_text, rapid_reason

    if ocr_text.strip() and not pdf_text_is_suspicious(ocr_text):
        if ocr_reason in {"OCR识别", "RapidOCR识别"}:
            logger.info("PDF text recovered with %s.", ocr_reason)
            return ocr_text, "已入库（OCR识别）"

    if best_reliable:
        logger.info("Using short but readable embedded PDF text from %s.", best_reliable[0])
        return best_reliable[1], "已入库"

    if ocr_text.strip() and pdf_text_is_suspicious(ocr_text):
        ocr_reason = "OCR结果仍异常"

    if candidates:
        logger.warning("PDF embedded text appears corrupted; refusing to index unreadable text.")
        return "", f"已保存（PDF文本疑似乱码，{ocr_reason}）"

    logger.warning("PDF has no usable embedded text; OCR result: %s.", ocr_reason)
    return "", f"已保存（PDF无可提取文字，{ocr_reason}）"
