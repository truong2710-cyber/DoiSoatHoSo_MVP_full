from __future__ import annotations

import contextlib
import hashlib
import json
import math
import multiprocessing as mp
import os
import pickle
import re
import subprocess
import sys
import threading
import time
import traceback
import unicodedata
from collections import Counter
from concurrent.futures import FIRST_COMPLETED, ProcessPoolExecutor, wait
from concurrent.futures.process import BrokenProcessPool
from dataclasses import dataclass, field
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from queue import Empty, Queue
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

# PyInstaller's windowed mode leaves standard streams unset. Some OCR and
# numerical libraries still write diagnostics to them.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

import fitz  # PyMuPDF: dùng cho cả đọc text lẫn render ảnh (đã bỏ pdfplumber)
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:  # tùy chọn: pip install psutil để đo CPU chính xác hơn
    import psutil
except ImportError:
    psutil = None


APP_NAME = "Đối soát hồ sơ"
OUTPUT_COLUMNS = [
    "Trạng thái đối soát", "Chi tiết sai lệch", "Độ tin cậy OCR",
    "Đường dẫn tài liệu", "Ghi chú kiểm tra thủ công",
]

# ----------------------------------------------------------------------------
# Cấu hình tốc độ (chỉnh ở đây)
# ----------------------------------------------------------------------------
OCR_DPI = 150            # trước đây 216. Nếu OCR sai nhiều thì thử 180.
MAX_IMAGE_SIDE = 1800    # cạnh dài tối đa của ảnh đưa vào OCR (px)
MAX_PAGES = 8            # số trang tối đa mỗi PDF
MAX_WORKERS = 6          # trần cứng; số worker thực tế tự thích ứng theo RAM/CPU
WORKER_RAM_GB = 1.5      # RAM ước tính mỗi tiến trình OCR
MIN_FREE_RAM_GB = 1.0    # luôn chừa lại ngần này RAM trống cho Windows
CPU_TARGET_PERCENT = 85  # CPU dưới mức này thì mới tăng thêm worker
USE_CACHE = True         # lưu kết quả quét để lần chạy sau không OCR lại
CACHE_VERSION = "v3-fitz150-earlyexit"   # đổi chuỗi này khi sửa logic trích xuất
CACHE_DIR = Path.home() / ".doi_soat_cache"
CACHE_SAVE_EVERY = 25    # lưu cache định kỳ để lỡ tắt app vẫn giữ được tiến độ

OCR_THREADS = os.cpu_count() or 2  # được ghi đè trong từng tiến trình con


class CancelledByUser(Exception):
    """Người dùng bấm Hủy."""


CANCEL_EVENT = threading.Event()


def resource_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def norm_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFD", str(value).strip().lower())
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.replace("đ", "d")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def norm_digits(value) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\D", "", str(value))


def ids_match(left: str, right: str) -> bool:
    if left == right:
        return bool(left)
    return (
        len(left) == 11 and len(right) == 12 and right.startswith("0") and right[1:] == left
    ) or (
        len(right) == 11 and len(left) == 12 and left.startswith("0") and left[1:] == right
    ) or (
        len(left) == 8 and len(right) == 9 and right.startswith("0") and right[1:] == left
    ) or (
        len(right) == 8 and len(left) == 9 and left.startswith("0") and left[1:] == right
    )


def is_valid_vietnamese_id(value: str, today: date | None = None) -> bool:
    if len(value) == 9 and value.isdigit():
        return True
    if len(value) != 12 or not value.isdigit():
        return False

    if not "001" <= value[:3] <= "096":
        return False

    century_code = int(value[3])
    century = 1900 if century_code < 2 else 2000 if century_code < 4 else 2100 + ((century_code - 4) // 2) * 100
    birth_year = century + int(value[4:6])
    return birth_year <= (today or date.today()).year


def vietnamese_id_birth_year(value: str) -> int | None:
    if len(value) != 12 or not value.isdigit():
        return None
    century_code = int(value[3])
    century = (
        1900
        if century_code < 2
        else 2000
        if century_code < 4
        else 2100 + ((century_code - 4) // 2) * 100
    )
    return century + int(value[4:6])


def is_plausible_driving_dob(value: str, today: date | None = None) -> bool:
    parsed = datetime.strptime(value, "%d/%m/%Y").date()
    current = today or date.today()
    age = current.year - parsed.year - ((current.month, current.day) < (parsed.month, parsed.day))
    return 18 <= age <= 100


def norm_date(value) -> str:
    if value is None or pd.isna(value) or value == "":
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.strftime("%d/%m/%Y")
    s = str(value).strip()
    if re.fullmatch(r"\d{8}", s):
        try:
            return datetime.strptime(s, "%Y%m%d").strftime("%d/%m/%Y")
        except ValueError:
            return ""
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    m = re.search(r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?!\d)", s)
    if m:
        d, mo, y = map(int, m.groups())
        if y < 100:
            y += 2000 if y < 30 else 1900
        try:
            return date(y, mo, d).strftime("%d/%m/%Y")
        except ValueError:
            return ""
    return ""


def alternate_norm_date(value) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        try:
            return date(value.year, value.day, value.month).strftime("%d/%m/%Y")
        except ValueError:
            return ""
    match = re.search(r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?!\d)", str(value).strip())
    if not match:
        return ""
    day, month, year = map(int, match.groups())
    if year < 100:
        year += 2000 if year < 30 else 1900
    try:
        return date(year, day, month).strftime("%d/%m/%Y")
    except ValueError:
        return ""


# ----------------------------------------------------------------------------
# OCR engine
# ----------------------------------------------------------------------------
_DEEPDOC_OCR = None
_DEEPDOC_OCR_LOCK = threading.Lock()


def get_deepdoc_ocr():
    global _DEEPDOC_OCR
    if _DEEPDOC_OCR is None:
        with _DEEPDOC_OCR_LOCK:
            if _DEEPDOC_OCR is None:
                engine_dir = resource_dir() / "deepdoc_vietocr"
                if not engine_dir.is_dir() and getattr(sys, "_MEIPASS", None):
                    engine_dir = resource_dir()

                required_models = [
                    engine_dir / "onnx" / "det.onnx",
                    engine_dir / "vietocr" / "weight" / "vgg_seq2seq.pth",
                ]
                missing = [str(path) for path in required_models if not path.is_file()]
                if missing:
                    raise FileNotFoundError(
                        "Thiếu model DeepDoc OCR: " + ", ".join(missing)
                    )

                if str(engine_dir) not in sys.path:
                    sys.path.insert(0, str(engine_dir))
                try:
                    from module.ocr import OCR
                except ModuleNotFoundError as exc:
                    raise RuntimeError(
                        f"Thiếu module Python '{exc.name}' khi nạp DeepDoc OCR. "
                        "Hãy kiểm tra bộ cài và các gói trong requirements.txt."
                    ) from exc
                engine = OCR()
                # Giới hạn số luồng PyTorch để nhiều tiến trình không tranh CPU nhau.
                try:
                    import torch

                    torch.set_num_threads(max(1, OCR_THREADS))
                except Exception:
                    pass
                _DEEPDOC_OCR = engine
    return _DEEPDOC_OCR


@contextlib.contextmanager
def _inference_context():
    """torch.inference_mode() nếu có torch: nhanh hơn và ít RAM hơn."""
    try:
        import torch

        with torch.inference_mode():
            yield
    except ImportError:
        yield


def open_folder(path: Path):
    """Mở thư mục kết quả trên Windows, macOS hoặc Linux."""
    folder = str(path.resolve())
    if sys.platform == "win32":
        os.startfile(folder)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])


def confirm_output_path(path: Path) -> Path | None:
    if not path.exists():
        return path
    overwrite = messagebox.askyesnocancel(
        APP_NAME,
        f"File kết quả đã tồn tại:\n{path}\n\n"
        "Chọn Yes để ghi đè, No để đổi tên file, hoặc Cancel để hủy.",
    )
    if overwrite is True:
        return path
    if overwrite is None:
        return None

    parent = path.parent
    initial_name = f"{path.stem}_moi{path.suffix}"
    while True:
        selected = filedialog.asksaveasfilename(
            initialdir=str(parent),
            initialfile=initial_name,
            defaultextension=path.suffix or ".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not selected:
            return None
        renamed_path = Path(selected)
        if not renamed_path.exists():
            return renamed_path
        messagebox.showwarning(
            APP_NAME,
            f"File này cũng đã tồn tại:\n{renamed_path}\nVui lòng chọn tên khác.",
        )
        parent = renamed_path.parent
        initial_name = renamed_path.name


def render_page(page: fitz.Page) -> np.ndarray:
    """Render trang bằng PyMuPDF (nhanh hơn pdfplumber nhiều), ảnh RGB."""
    rect = page.rect
    scale = OCR_DPI / 72.0
    longest = max(rect.width, rect.height) * scale
    if longest > MAX_IMAGE_SIDE:
        scale *= MAX_IMAGE_SIDE / longest
    pix = page.get_pixmap(
        matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False
    )
    image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, 3
    )
    return image.copy()


def is_blank_image(image: np.ndarray) -> bool:
    """Trang trắng/gần trắng: bỏ qua OCR để tiết kiệm vài giây."""
    return float(image[::8, ::8].std()) < 2.0


def ocr_page(
    page: fitz.Page,
    rotation_degrees: int = 0,
) -> tuple[
    str,
    float,
    list[tuple[str, float]],
    list[tuple[str, float, tuple[float, float, float, float]]],
]:
    if rotation_degrees not in (0, 90, -90):
        raise ValueError("rotation_degrees must be 0, 90, or -90")
    image = render_page(page)
    if is_blank_image(image):
        return "", 0.0, [], []
    if rotation_degrees:
        image = np.rot90(image, k=-rotation_degrees // 90).copy()
    with _inference_context():
        results = get_deepdoc_ocr()(image)
    recognized = []
    layout_lines = []
    for box, (text, score) in (results or []):
        if not text or not text.strip():
            continue
        recognized.append((text.strip(), float(score)))
        points = np.asarray(box, dtype=float)
        if points.size < 8:
            continue
        points = points.reshape(-1, 2)
        points[:, 0] /= image.shape[1]
        points[:, 1] /= image.shape[0]
        if rotation_degrees == 90:
            points = np.column_stack((points[:, 1], 1 - points[:, 0]))
        elif rotation_degrees == -90:
            points = np.column_stack((1 - points[:, 1], points[:, 0]))
        bounds = tuple(
            float(value)
            for value in (
                points[:, 0].min(),
                points[:, 1].min(),
                points[:, 0].max(),
                points[:, 1].max(),
            )
        )
        layout_lines.append((text.strip(), float(score) * 100, bounds))
    recognized_text = recognized
    text = " ".join(word for word, _ in recognized_text)
    character_count = sum(len(word) for word, _ in recognized_text)
    confidence = (
        sum(len(word) * score for word, score in recognized_text) / character_count
        if character_count
        else 0.0
    )
    return text, confidence * 100, [
        (word, score * 100) for word, score in recognized_text
    ], layout_lines


_BIRTH_LABEL_PATTERN = re.compile(
    r"\b(?:sinh ngay|ngay(?: thang nam)? (?:sinh|tinh)|nam sinh|date of birth|date of binh|dob)\b"
)
_OVERLONG_ID_PATTERN = re.compile(r"(?<!\d)(?:\d[ .-]?){12,14}\d(?!\d)")
_ID_PATTERN = re.compile(r"(?<!\d)(?:\d[ .-]?){8,11}\d(?!\d)")


def _birth_date_labels(
    lines: list[tuple[str, float | None, tuple[float, float, float, float]]],
    min_confidence: float = 50.0,
) -> list[tuple[str, float | None, tuple[float, float, float, float]]]:
    return [
        line
        for line in lines
        if (line[1] is None or line[1] >= min_confidence)
        and _BIRTH_LABEL_PATTERN.search(norm_text(line[0]))
    ]


def _orientation_quality(
    lines: list[tuple[str, float | None]],
    layout: list[
        tuple[str, float | None, tuple[float, float, float, float]]
    ],
    context_labels: list[
        tuple[str, float | None, tuple[float, float, float, float]]
    ] | None = None,
) -> tuple[
    tuple[int, float, float, int],
    list[tuple[str, float | None, tuple[float, float, float, float]]],
]:
    effective_layout = list(layout)
    if not _birth_date_labels(effective_layout) and context_labels:
        effective_layout.extend(context_labels)
    votes, _ = extract_spatial_birth_date_votes(effective_layout)
    id_candidates = extract_id_candidates("", lines, recover_overlong=True)
    _, id_confidence = select_id_candidate(id_candidates)
    vote_weight = sum(
        0.5 + 0.5 * (score if score is not None else 50.0) / 100
        for _, score in votes
    )
    quality = (
        len(votes),
        vote_weight,
        id_confidence if id_confidence is not None else 0.0,
        len(id_candidates),
    )
    return quality, effective_layout


def _retry_ocr_orientation(
    lines: list[tuple[str, float | None]],
    confidence: float,
    layout: list[
        tuple[str, float | None, tuple[float, float, float, float]]
    ],
    document_has_birth_date_label: bool = False,
    document_has_birth_date: bool = False,
) -> bool:
    labels = _birth_date_labels(layout)
    votes, _ = extract_spatial_birth_date_votes(layout)
    high_confidence_vote = any(
        score is not None and score >= 80.0 for _, score in votes
    )
    candidates = extract_id_candidates("", lines, recover_overlong=True)
    id_candidate, id_confidence = select_id_candidate(candidates)
    weak_id = id_candidate is not None and (
        id_confidence is None or id_confidence < 80.0
    )
    return (
        (
            (
                bool(labels)
                or (document_has_birth_date_label and not document_has_birth_date)
            )
            and not high_confidence_vote
        )
        or weak_id
        or (confidence < 65.0 and not high_confidence_vote)
    )


def _has_confident_id_and_dob(
    ocr_lines: list[tuple[str, float | None]],
    layouts: list,
) -> bool:
    """Đã có CCCD 12 số (>=80%) và ngày sinh (>=80%) từ OCR -> khỏi OCR thêm trang."""
    scored = [(t, s) for t, s in ocr_lines if s is not None]
    if not scored:
        return False
    candidates = extract_id_candidates("", scored, recover_overlong=True)
    selected, conf = select_id_candidate(candidates, min_confidence=60.0)
    if not selected or len(selected) != 12 or conf is None or conf < 80.0:
        return False
    for layout in layouts:
        if not layout:
            continue
        votes, _ = extract_spatial_birth_date_votes(layout)
        if any(score is not None and score >= 80.0 for _, score in votes):
            return True
    return False


@dataclass
class DocumentInfo:
    path: Path
    text: str
    ids: set[str]
    dates: set[str]
    birth_dates: set[str]
    name_hint: str
    confidence: float | None
    used_ocr: bool
    error: str = ""
    text_lines: list[tuple[str, float | None]] = field(default_factory=list)
    layout_pages: list[
        list[tuple[str, float | None, tuple[float, float, float, float]]] | None
    ] = field(default_factory=list)
    id_confidence: float | None = None
    birth_date_confidence: float | None = None
    birth_date_out_of_range: bool = False
    birth_date_ambiguous: bool = False
    birth_date_wide_search: bool = False
    id_from_pdf: bool = False
    filename_ids: set[str] = field(default_factory=set)


def inspect_pdf(path: Path) -> DocumentInfo:
    chunks, ocr_scores, text_lines, text_pages, layout_pages, used_ocr, error = (
        [], [], [], [], [], False, ""
    )
    document_has_birth_date_label = False
    document_has_birth_date = False
    skip_further_ocr = False   # đã đủ CCCD + ngày sinh tin cậy -> bỏ OCR các trang sau
    good_rotation: int | None = None
    ocr_only_lines: list[tuple[str, float | None]] = []
    try:
        with fitz.open(path) as doc:
            for page_index in range(min(len(doc), MAX_PAGES)):
                page = doc[page_index]
                text = page.get_text("text").strip()
                if len(re.sub(r"\s", "", text)) >= 40:
                    chunks.append(text)
                    page_lines = [
                        (line.strip(), None)
                        for line in text.splitlines()
                        if line.strip()
                    ]
                    page_layout = None
                else:
                    if skip_further_ocr:
                        continue
                    used_ocr = True
                    ocr_text, conf, page_lines, page_layout = ocr_page(page)
                    base_page_lines = list(page_lines)
                    document_has_birth_date_label |= bool(
                        _birth_date_labels(page_layout)
                    )
                    if page_layout and _retry_ocr_orientation(
                        page_lines,
                        conf,
                        page_layout,
                        document_has_birth_date_label,
                        document_has_birth_date,
                    ):
                        base_labels = _birth_date_labels(page_layout)
                        best_quality, _ = _orientation_quality(
                            page_lines, page_layout
                        )
                        # Nếu trang trước đã phải xoay thì thử đúng chiều đó trước.
                        rotations = (-90, 90) if good_rotation == -90 else (90, -90)
                        for rotation in rotations:
                            (
                                rotated_text,
                                rotated_conf,
                                rotated_lines,
                                rotated_layout,
                            ) = ocr_page(page, rotation_degrees=rotation)
                            quality, effective_layout = _orientation_quality(
                                rotated_lines,
                                rotated_layout,
                                context_labels=base_labels,
                            )
                            if quality <= best_quality:
                                continue
                            best_quality = quality
                            good_rotation = rotation
                            ocr_text = rotated_text
                            conf = rotated_conf
                            page_lines = list(rotated_lines)
                            if not _birth_date_labels(rotated_layout):
                                page_lines.extend(
                                    (t, s) for t, s, _ in base_labels
                                )
                            rotated_ids = extract_id_candidates(
                                "", page_lines, recover_overlong=True
                            )
                            if not rotated_ids:
                                page_lines.extend(
                                    line
                                    for line in base_page_lines
                                    if extract_id_candidates(
                                        "", [line], recover_overlong=True
                                    )
                                )
                            page_layout = effective_layout
                            # Kết quả đã đủ tốt -> không cần thử hướng còn lại.
                            if not _retry_ocr_orientation(
                                page_lines,
                                conf,
                                page_layout,
                                document_has_birth_date_label,
                                document_has_birth_date,
                            ):
                                break
                    chunks.append(ocr_text)
                    ocr_scores.append((conf, max(len(re.sub(r"\s", "", ocr_text)), 1)))
                    ocr_only_lines.extend(page_lines)
                document_has_birth_date |= any(
                    extract_ocr_birth_dates(line)
                    for line, _ in page_lines
                )
                text_lines.extend(page_lines)
                text_pages.append(page_lines)
                layout_pages.append(page_layout or None)
                if used_ocr and not skip_further_ocr:
                    skip_further_ocr = _has_confident_id_and_dob(
                        ocr_only_lines, layout_pages
                    )
    except Exception as exc:
        error = str(exc)
    raw = "\n".join(chunks)
    filename_candidates = extract_id_candidates(path.stem, [])
    filename_ids = {candidate for candidate, _ in filename_candidates}
    pdf_candidates = extract_id_candidates("", text_lines)
    ocr_lines = [(text, score) for text, score in text_lines if score is not None]
    recovered_candidates = extract_id_candidates(
        "", ocr_lines, recover_overlong=True
    )
    pdf_candidates = [
        candidate
        for candidate in (*pdf_candidates, *recovered_candidates)
        if len(candidate[0]) == 12
    ]
    corroborated_candidates = [
        candidate for candidate in pdf_candidates if candidate[0] in filename_ids
    ]
    if corroborated_candidates:
        pdf_candidates = corroborated_candidates
    selected_id, id_confidence = select_id_candidate(
        pdf_candidates, min_confidence=60.0
    )
    id_from_pdf = selected_id is not None
    if selected_id is None:
        selected_id, id_confidence = select_id_candidate(filename_candidates)
    inferred_birth_year = (
        vietnamese_id_birth_year(selected_id)
        if id_from_pdf and id_confidence is not None and id_confidence >= 80.0
        else None
    )
    ids = {selected_id} if selected_id else set()
    dates = set()
    for d, m, y in re.findall(r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?!\d)", raw):
        value = norm_date(f"{d}/{m}/{y}")
        if value:
            dates.add(value)
    (
        birth_date,
        birth_date_confidence,
        birth_date_out_of_range,
        birth_date_ambiguous,
        birth_date_wide_search,
    ) = extract_birth_date_from_pages(
        text_pages,
        layout_pages=layout_pages,
        inferred_birth_year=inferred_birth_year,
    )
    birth_dates = {birth_date} if birth_date else set()
    name_hint = re.sub(r"\b\d[\d ._-]*\b", " ", path.stem)
    name_hint = norm_text(name_hint)
    confidence = (
        sum(score * weight for score, weight in ocr_scores)
        / sum(weight for _, weight in ocr_scores)
        if used_ocr and ocr_scores
        else None if used_ocr else 100.0
    )
    return DocumentInfo(
        path=path,
        text=raw,
        ids=ids,
        dates=dates,
        birth_dates=birth_dates,
        name_hint=name_hint,
        confidence=confidence,
        used_ocr=used_ocr,
        error=error,
        text_lines=text_lines,
        layout_pages=layout_pages,
        id_confidence=id_confidence,
        birth_date_confidence=birth_date_confidence,
        birth_date_out_of_range=birth_date_out_of_range,
        birth_date_ambiguous=birth_date_ambiguous,
        birth_date_wide_search=birth_date_wide_search,
        id_from_pdf=id_from_pdf,
        filename_ids={candidate for candidate, _ in filename_candidates},
    )


def extract_id_candidates(
    filename: str,
    lines: list[tuple[str, float | None]],
    recover_overlong: bool = False,
) -> list[tuple[str, float | None]]:
    candidates: list[tuple[str, float | None]] = []
    seen: dict[str, int] = {}

    def add(candidate: str, score: float | None) -> None:
        digits = norm_digits(candidate)
        if not is_valid_vietnamese_id(digits):
            return
        if digits not in seen:
            seen[digits] = len(candidates)
            candidates.append((digits, score))
        elif score is not None:
            index = seen[digits]
            old_score = candidates[index][1]
            if old_score is None or score > old_score:
                candidates[index] = (digits, score)

    for match in _ID_PATTERN.finditer(filename):
        add(match.group(), None)
    for text, score in lines:
        for match in _ID_PATTERN.finditer(text):
            add(match.group(), score)
        if recover_overlong:
            for match in _OVERLONG_ID_PATTERN.finditer(text):
                digits = norm_digits(match.group())
                for start in range(len(digits) - 11):
                    add(digits[start:start + 12], score)

    for index in range(len(lines) - 1):
        first, first_score = lines[index]
        second, second_score = lines[index + 1]
        tail = re.search(r"(\d[\d .-]*)$", first.strip())
        head = re.match(r"(\d[\d .-]*)", second.strip())
        if not tail or not head:
            continue
        first_digits = norm_digits(tail.group(1))
        second_digits = norm_digits(head.group(1))
        if len(first_digits) + len(second_digits) not in (9, 12):
            continue
        score = (
            (first_score + second_score) / 2
            if first_score is not None and second_score is not None
            else None
        )
        add(first_digits + second_digits, score)
    return candidates


def select_id_candidate(
    candidates: list[tuple[str, float | None]],
    min_confidence: float = 60.0,
) -> tuple[str | None, float | None]:
    if not candidates:
        return None, None
    eligible = [
        (index, candidate, score)
        for index, (candidate, score) in enumerate(candidates)
        if score is None or score >= min_confidence
    ]
    if not eligible:
        return None, None
    candidate, score = max(
        eligible,
        key=lambda item: (
            len(item[1]) == 12,
            item[2] is not None,
            item[2] if item[2] is not None else 0.0,
            -item[0],
        ),
    )[1:3]
    return candidate, score


_SLASH_BIRTH_DATE_PATTERN = re.compile(
    r"(?<![\d/])(?P<day>\d{1,2})\s*/\s*(?P<month>\d{1,2})\s*/\s*"
    r"(?P<year>\d(?:[\s.-]?\d){0,5})(?![\d/-])"
)
_OTHER_BIRTH_DATE_PATTERN = re.compile(
    r"(?<![\d./-])(?P<day>\d{1,2})\s*[./-]\s*"
    r"(?P<month>\d{1,2})\s*[./-]\s*"
    r"(?P<year>\d(?:[\s.-]?\d){0,5})(?![\d/-])"
)
_COMPACT_SLASH_BIRTH_DATE_PATTERN = re.compile(
    r"(?<![\d/])(?P<day>\d{1,2})\s*/\s*(?P<packed>\d{6,8})(?!\d)"
)
_PARTIAL_SLASH_BIRTH_DATE_PATTERN = re.compile(
    r"(?<![\d/])(?P<day>\d{1,2})\s*/\s*(?P<month>\d{1,2})(?![\d/])"
)


def _normalize_ocr_birth_date(
    day_text: str,
    month_text: str,
    year_text: str,
    inferred_birth_year: int | None,
) -> str:
    year_digits = re.sub(r"\D", "", year_text)
    year_is_uncertain = (
        len(year_digits) != 4 or bool(re.search(r"[.\s-]", year_text))
    )
    if year_is_uncertain and inferred_birth_year is not None:
        year = inferred_birth_year
    elif len(year_digits) == 4:
        year = int(year_digits)
    elif len(year_digits) == 2:
        short_year = int(year_digits)
        year = 2000 + short_year if short_year < 30 else 1900 + short_year
    else:
        return ""
    try:
        return date(int(year), int(month_text), int(day_text)).strftime("%d/%m/%Y")
    except ValueError:
        return ""


def extract_ocr_birth_dates(
    text: str,
    inferred_birth_year: int | None = None,
) -> list[str]:
    def normalize_matches(matches: list[re.Match[str]]) -> list[str]:
        values = [
            _normalize_ocr_birth_date(
                match.group("day"),
                match.group("month"),
                match.group("year"),
                inferred_birth_year,
            )
            for match in matches
        ]
        return [value for value in values if value]

    values = normalize_matches(list(_SLASH_BIRTH_DATE_PATTERN.finditer(text)))
    if values:
        return values
    values = normalize_matches(list(_OTHER_BIRTH_DATE_PATTERN.finditer(text)))
    if values:
        return values

    compact_match = _COMPACT_SLASH_BIRTH_DATE_PATTERN.search(text)
    if compact_match:
        packed = compact_match.group("packed")
        month_text, year_text = packed[:2], packed[2:]
        value = _normalize_ocr_birth_date(
            compact_match.group("day"),
            month_text,
            year_text,
            inferred_birth_year,
        )
        return [value] if value else []

    if inferred_birth_year is not None:
        partial_match = _PARTIAL_SLASH_BIRTH_DATE_PATTERN.search(text)
        if partial_match:
            value = _normalize_ocr_birth_date(
                partial_match.group("day"),
                partial_match.group("month"),
                str(inferred_birth_year),
                inferred_birth_year,
            )
            return [value] if value else []
    return []


def extract_birth_date_candidate(
    lines: list[tuple[str, float | None]],
    max_line_distance: int = 12,
    min_confidence: float = 50.0,
    today: date | None = None,
    near_line_distance: int = 2,
    inferred_birth_year: int | None = None,
) -> tuple[str | None, float | None, bool, bool, bool]:
    labels = [
        index
        for index, (text, confidence) in enumerate(lines)
        if (confidence is None or confidence >= min_confidence)
        and _BIRTH_LABEL_PATTERN.search(norm_text(text))
    ]
    date_candidates: list[tuple[str, float | None, int, int]] = []

    def add_dates(text: str, start_line: int, end_line: int, score: float | None) -> None:
        if score is not None and score < min_confidence:
            return
        for value in extract_ocr_birth_dates(text, inferred_birth_year):
            date_candidates.append((value, score, start_line, end_line))

    for index, (text, confidence) in enumerate(lines):
        add_dates(text, index, index, confidence)
        if index + 1 < len(lines):
            next_text, next_confidence = lines[index + 1]
            joined = text + " " + next_text
            individual_values = set(
                extract_ocr_birth_dates(text, inferred_birth_year)
                + extract_ocr_birth_dates(next_text, inferred_birth_year)
            )
            for value in extract_ocr_birth_dates(joined, inferred_birth_year):
                if value not in individual_values:
                    score = (
                        (confidence + next_confidence) / 2
                        if confidence is not None and next_confidence is not None
                        else None
                    )
                    date_candidates.append((value, score, index, index + 1))

    eligible: list[tuple[str, float | None, int, int, int]] = []
    out_of_range = False
    for value, score, start_line, end_line in date_candidates:
        distances = [
            0 if start_line <= label <= end_line
            else start_line - label if label < start_line
            else label - end_line
            for label in labels
        ]
        distance = min(distances, default=max_line_distance + 1)
        if distance > max_line_distance:
            continue
        if not is_plausible_driving_dob(value, today):
            out_of_range = True
            continue
        eligible.append((value, score, distance, start_line, end_line))

    if not eligible:
        return None, None, out_of_range, False, False
    nearby = [
        candidate for candidate in eligible
        if candidate[2] <= near_line_distance
    ]
    used_wide_search = not nearby
    considered = nearby or eligible
    nearest_distance = min(candidate[2] for candidate in considered)
    nearest = [
        candidate for candidate in considered
        if candidate[2] == nearest_distance
    ]
    best_score = max(
        (
            candidate[1] is not None,
            candidate[1] if candidate[1] is not None else 0.0,
        )
        for candidate in nearest
    )
    best_dates = {
        candidate[0]
        for candidate in nearest
        if (
            candidate[1] is not None,
            candidate[1] if candidate[1] is not None else 0.0,
        ) == best_score
    }
    if len(best_dates) > 1:
        return None, None, out_of_range, True, used_wide_search
    _, (value, score, _, _, _) = max(
        enumerate(nearest),
        key=lambda item: (
            item[1][1] is not None,
            item[1][1] if item[1][1] is not None else 0.0,
            -item[1][3],
            -item[0],
        ),
    )
    return value, score, out_of_range, False, used_wide_search


def extract_spatial_birth_date_votes(
    layout_lines: list[
        tuple[str, float | None, tuple[float, float, float, float]]
    ],
    min_confidence: float = 50.0,
    max_normalized_distance: float = 0.04,
    today: date | None = None,
    inferred_birth_year: int | None = None,
) -> tuple[list[tuple[str, float | None]], bool]:
    labels = _birth_date_labels(layout_lines, min_confidence)
    if not labels:
        # Không có nhãn "ngày sinh" thì không thể có phiếu bầu -> khỏi parse ngày.
        return [], False
    dates: list[
        tuple[str, float | None, tuple[float, float, float, float], bool]
    ] = []
    out_of_range = False
    for text, score, bounds in layout_lines:
        for value in extract_ocr_birth_dates(text, inferred_birth_year):
            plausible = is_plausible_driving_dob(value, today)
            dates.append((value, score, bounds, plausible))

    votes: list[tuple[str, float | None]] = []
    for _, _, label_bounds in labels:
        lx0, ly0, lx1, ly1 = label_bounds
        nearby_dates = []
        for value, score, date_bounds, plausible in dates:
            if score is not None and score < min_confidence:
                continue
            dx0, dy0, dx1, dy1 = date_bounds
            dx = max(0.0, lx0 - dx1, dx0 - lx1)
            dy = max(0.0, ly0 - dy1, dy0 - ly1)
            distance = min(dx, dy)
            if distance > max_normalized_distance:
                continue
            if not plausible:
                out_of_range = True
                continue
            nearby_dates.append((distance, value, score))
        if not nearby_dates:
            continue
        nearest_distance = min(candidate[0] for candidate in nearby_dates)
        nearest = [
            candidate for candidate in nearby_dates
            if math.isclose(candidate[0], nearest_distance, abs_tol=0.002)
        ]
        nearest_values = {candidate[1] for candidate in nearest}
        if len(nearest_values) != 1:
            continue
        selected = max(
            nearest, key=lambda candidate: (
                candidate[2] is not None,
                candidate[2] if candidate[2] is not None else 0.0,
            )
        )
        votes.append((selected[1], selected[2]))
    return votes, out_of_range


def extract_birth_date_from_pages(
    pages: list[list[tuple[str, float | None]]],
    max_line_distance: int = 12,
    min_confidence: float = 50.0,
    today: date | None = None,
    layout_pages: list[
        list[tuple[str, float | None, tuple[float, float, float, float]]] | None
    ] | None = None,
    inferred_birth_year: int | None = None,
) -> tuple[str | None, float | None, bool, bool, bool]:
    if layout_pages is not None:
        votes: list[tuple[str, float | None]] = []
        out_of_range = False
        for index, lines in enumerate(pages):
            layout = layout_pages[index] if index < len(layout_pages) else None
            if layout is None:
                candidate = extract_birth_date_candidate(
                    lines,
                    max_line_distance,
                    min_confidence,
                    today,
                    inferred_birth_year=inferred_birth_year,
                )
                out_of_range |= candidate[2]
                if candidate[0] is not None:
                    votes.append((candidate[0], candidate[1]))
            else:
                page_votes, page_out_of_range = extract_spatial_birth_date_votes(
                    layout,
                    min_confidence=min_confidence,
                    today=today,
                    inferred_birth_year=inferred_birth_year,
                )
                votes.extend(page_votes)
                out_of_range |= page_out_of_range
        if not votes:
            return None, None, out_of_range, False, False

        vote_weights: dict[str, float] = {}
        scores: dict[str, list[float]] = {}
        for value, confidence in votes:
            effective_confidence = (
                confidence if confidence is not None else min_confidence
            )
            weight = 0.5 + 0.5 * effective_confidence / 100
            vote_weights[value] = vote_weights.get(value, 0.0) + weight
            if confidence is not None:
                scores.setdefault(value, []).append(confidence)

        ranked = sorted(vote_weights.items(), key=lambda item: item[1], reverse=True)
        if (
            len(ranked) > 1
            and ranked[0][1] - ranked[1][1] <= 0.004 + 1e-12
        ):
            return None, None, out_of_range, True, False
        winner = ranked[0][0]
        winning_scores = scores.get(winner, [])
        mean_confidence = (
            sum(winning_scores) / len(winning_scores) if winning_scores else None
        )
        return winner, mean_confidence, out_of_range, False, False

    candidates = [
        extract_birth_date_candidate(
            lines,
            max_line_distance,
            min_confidence,
            today,
            inferred_birth_year=inferred_birth_year,
        )
        for lines in pages
    ]
    eligible = [candidate for candidate in candidates if candidate[0] is not None]
    if not eligible:
        return (
            None,
            None,
            any(candidate[2] for candidate in candidates),
            any(candidate[3] for candidate in candidates),
            False,
        )

    nearby = [candidate for candidate in eligible if not candidate[4]]
    considered = nearby or eligible
    best_score = max(
        (
            candidate[1] is not None,
            candidate[1] if candidate[1] is not None else 0.0,
        )
        for candidate in considered
    )
    best = [
        candidate
        for candidate in considered
        if (
            candidate[1] is not None,
            candidate[1] if candidate[1] is not None else 0.0,
        ) == best_score
    ]
    best_dates = {candidate[0] for candidate in best}
    if len(best_dates) > 1:
        return None, None, False, True, not nearby
    value, score, out_of_range, ambiguous, wide_search = best[0]
    return value, score, out_of_range, ambiguous, wide_search


def extract_birth_dates(
    lines: list[tuple[str, float | None]],
    max_line_distance: int = 12,
    min_confidence: float = 50.0,
    today: date | None = None,
) -> set[str]:
    value, _, _, _, _ = extract_birth_date_candidate(
        lines, max_line_distance, min_confidence, today
    )
    return {value} if value else set()


ALIASES = {
    "name": ["ho va ten", "ho ten", "ten khach hang", "ten nguoi", "customer name"],
    "id": ["so cccd", "cccd", "cmnd", "so cmnd", "so chung minh", "chung minh nhan dan", "can cuoc", "identity"],
    "dob": ["ngay sinh", "nam sinh", "date of birth", "dob"],
}


def detect_columns(columns) -> dict[str, str | None]:
    normalized = {c: norm_text(c) for c in columns}
    result = {}
    for field_name, aliases in ALIASES.items():
        result[field_name] = next((c for c, n in normalized.items() if n in aliases), None)
        if not result[field_name]:
            result[field_name] = next((c for c, n in normalized.items() if any(a in n for a in aliases)), None)
    return result


def match_document(
    name: str,
    citizen_id: str,
    docs: list[DocumentInfo],
    allow_name_fallback: bool = True,
) -> tuple[DocumentInfo | None, float]:
    if citizen_id:
        exact_pdf_ids = [
            d for d in docs
            if d.id_from_pdf and any(ids_match(citizen_id, candidate) for candidate in d.ids)
        ]
        if exact_pdf_ids:
            return exact_pdf_ids[0], 1.0
        exact_filename_ids = [
            d for d in docs
            if any(ids_match(citizen_id, candidate) for candidate in d.filename_ids)
        ]
        if exact_filename_ids:
            normalized_name = norm_text(name)
            best_match = max(
                exact_filename_ids,
                key=lambda doc: SequenceMatcher(
                    None, normalized_name, doc.name_hint
                ).ratio(),
            )
            return best_match, 1.0
    if not allow_name_fallback:
        return None, 0.0
    n = norm_text(name)
    if not n:
        return None, 0.0
    # Nếu cả Excel và PDF đều có CCCD nhưng khác nhau thì tuyệt đối không ghép
    # chỉ vì tên gần giống: đây là lỗi nguy hiểm nhất khi danh sách dài.
    candidates = [
        d for d in docs
        if d.id_from_pdf
        and (not citizen_id or not d.ids or any(ids_match(citizen_id, candidate) for candidate in d.ids))
    ]
    ranked = sorted(((SequenceMatcher(None, n, d.name_hint).ratio(), d) for d in candidates), reverse=True, key=lambda x: x[0])
    return (ranked[0][1], ranked[0][0]) if ranked and ranked[0][0] >= 0.67 else (None, ranked[0][0] if ranked else 0.0)


def _ratio_at_least(a: str, b: str, threshold: float) -> float:
    """SequenceMatcher.ratio() nhưng loại nhanh các cặp chắc chắn dưới ngưỡng.

    real_quick_ratio / quick_ratio là cận trên của ratio(), nên kết quả
    (so với ngưỡng) giống hệt việc gọi ratio() trực tiếp nhưng nhanh hơn nhiều.
    """
    matcher = SequenceMatcher(None, a, b)
    if matcher.real_quick_ratio() < threshold or matcher.quick_ratio() < threshold:
        return 0.0
    return matcher.ratio()


def filename_similarity(
    name: str, citizen_id: str, doc: DocumentInfo
) -> tuple[float, float, float] | None:
    normalized_name = norm_text(name)
    normalized_id = norm_digits(citizen_id)
    if not normalized_name or not normalized_id or not doc.name_hint or not doc.filename_ids:
        return None

    name_score = _ratio_at_least(normalized_name, doc.name_hint, 0.85)
    if name_score < 0.85:
        return None
    id_score = max(
        1.0 if ids_match(normalized_id, filename_id)
        else _ratio_at_least(normalized_id, filename_id, 0.85)
        for filename_id in doc.filename_ids
    )
    if id_score < 0.85:
        return None
    return (name_score + id_score) / 2, name_score, id_score


def assign_filename_similarity_matches(
    df: pd.DataFrame,
    name_column: str,
    id_column: str,
    docs: list[DocumentInfo],
) -> dict[int, tuple[DocumentInfo, float, float, float]]:
    exact_rows: set[int] = set()
    exact_docs: set[Path] = set()
    row_values = []
    for row_index, (_, row) in enumerate(df.iterrows()):
        name = str(row.get(name_column, "") or "")
        citizen_id = norm_digits(row.get(id_column, ""))
        row_values.append((name, citizen_id))
        if not citizen_id:
            continue
        for doc in docs:
            if (
                doc.id_from_pdf
                and any(ids_match(citizen_id, candidate) for candidate in doc.ids)
            ) or any(ids_match(citizen_id, candidate) for candidate in doc.filename_ids):
                exact_rows.add(row_index)
                exact_docs.add(doc.path)

    candidates = []
    remaining_docs = [doc for doc in docs if doc.path not in exact_docs]
    for row_index, (name, citizen_id) in enumerate(row_values):
        if row_index in exact_rows:
            continue
        for doc in remaining_docs:
            scores = filename_similarity(name, citizen_id, doc)
            if scores:
                combined, name_score, id_score = scores
                candidates.append(
                    (combined, name_score, id_score, row_index, doc)
                )

    assignments = {}
    assigned_docs: set[Path] = set()
    for combined, name_score, id_score, row_index, doc in sorted(
        candidates, key=lambda item: (-item[0], -item[1], -item[2], item[3], str(item[4].path))
    ):
        if row_index in assignments or doc.path in assigned_docs:
            continue
        assignments[row_index] = (doc, combined, name_score, id_score)
        assigned_docs.add(doc.path)
    return assignments


# ----------------------------------------------------------------------------
# Cache + quét PDF song song
# ----------------------------------------------------------------------------
def _cache_file(pdf_dir: Path) -> Path:
    key = hashlib.md5(str(pdf_dir.resolve()).encode("utf-8")).hexdigest()[:12]
    return CACHE_DIR / f"scan_{key}.pkl"


def _load_cache(pdf_dir: Path) -> dict:
    if not USE_CACHE:
        return {}
    try:
        with open(_cache_file(pdf_dir), "rb") as handle:
            payload = pickle.load(handle)
        if payload.get("version") == CACHE_VERSION:
            return payload.get("items", {})
    except Exception:
        pass
    return {}


def _save_cache(pdf_dir: Path, items: dict) -> None:
    if not USE_CACHE:
        return
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        target = _cache_file(pdf_dir)
        temp = target.with_suffix(".tmp")
        with open(temp, "wb") as handle:
            pickle.dump({"version": CACHE_VERSION, "items": items}, handle)
        os.replace(temp, target)
    except Exception:
        pass  # cache lỗi không được làm hỏng quá trình đối soát


def _file_signature(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return stat.st_size, stat.st_mtime_ns


def _init_worker(threads: int) -> None:
    """Chạy một lần trong mỗi tiến trình con: giới hạn số luồng để không tranh CPU."""
    global OCR_THREADS
    OCR_THREADS = max(1, threads)
    for name in ("OMP_NUM_THREADS", "MKL_NUM_THREADS", "OPENBLAS_NUM_THREADS", "NUMEXPR_NUM_THREADS"):
        os.environ[name] = str(OCR_THREADS)


def _error_info(path: Path, exc: Exception) -> DocumentInfo:
    return DocumentInfo(
        path=path, text="", ids=set(), dates=set(), birth_dates=set(),
        name_hint=norm_text(re.sub(r"\b\d[\d ._-]*\b", " ", path.stem)),
        confidence=None, used_ocr=False, error=str(exc),
        filename_ids={c for c, _ in extract_id_candidates(path.stem, [])},
    )


def available_ram_gb() -> float | None:
    """RAM còn trống (GB). Trả None nếu không đo được."""
    if psutil is not None:
        return psutil.virtual_memory().available / 2**30
    if sys.platform == "win32":
        try:
            import ctypes

            class MemoryStatus(ctypes.Structure):
                _fields_ = [
                    ("dwLength", ctypes.c_ulong),
                    ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong),
                    ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong),
                    ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong),
                    ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("sullAvailExtendedVirtual", ctypes.c_ulonglong),
                ]

            stat = MemoryStatus()
            stat.dwLength = ctypes.sizeof(MemoryStatus)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat))
            return stat.ullAvailPhys / 2**30
        except Exception:
            return None
    try:
        with open("/proc/meminfo") as handle:
            for line in handle:
                if line.startswith("MemAvailable:"):
                    return int(line.split()[1]) / 2**20
    except OSError:
        pass
    return None


def cpu_percent() -> float | None:
    if psutil is None:
        return None
    return psutil.cpu_percent(interval=None)


def choose_worker_plan(pending: int) -> tuple[int, int]:
    """Số worker tối đa = min(số nhân - 1, số worker mà RAM trống chịu được)."""
    cores = os.cpu_count() or 2
    by_cpu = max(1, cores - 1)
    ram = available_ram_gb()
    by_ram = (
        max(1, int((ram - MIN_FREE_RAM_GB) // WORKER_RAM_GB)) if ram is not None else 2
    )
    workers = max(1, min(by_cpu, by_ram, MAX_WORKERS, pending))
    threads = max(1, round(cores / workers))
    return workers, threads


def adapt_concurrency(limit: int, workers: int) -> int:
    """Tăng/giảm số PDF xử lý đồng thời theo RAM và CPU đang rảnh."""
    ram = available_ram_gb()
    if ram is not None and ram < MIN_FREE_RAM_GB:
        return max(1, limit - 1)          # sắp thiếu RAM -> giảm
    cpu = cpu_percent()
    has_ram_for_more = ram is None or ram > WORKER_RAM_GB + MIN_FREE_RAM_GB
    cpu_has_room = cpu is None or cpu < CPU_TARGET_PERCENT
    if limit < workers and has_ram_for_more and cpu_has_room:
        return limit + 1                  # còn dư RAM và CPU -> tăng
    return limit


def scan_pdfs(pdfs: list[Path], pdf_dir: Path, progress, on_scan) -> list[DocumentInfo]:
    total = len(pdfs)
    docs: list[DocumentInfo | None] = [None] * total
    cache = _load_cache(pdf_dir)
    done = 0
    pending: list[int] = []

    for index, path in enumerate(pdfs):
        hit = None
        try:
            size, mtime = _file_signature(path)
            entry = cache.get(str(path.resolve()))
            if entry and entry[0] == size and entry[1] == mtime:
                hit = entry[2]
        except OSError:
            pass
        if hit is not None:
            hit.path = path
            docs[index] = hit
            done += 1
            progress("scan", done, total, f"Đã có trong cache: {path.name}")
            if on_scan:
                on_scan(index + 1, total, path, "start", None)
                on_scan(index + 1, total, path, "done", hit)
        else:
            pending.append(index)

    def finish(index: int, info: DocumentInfo) -> None:
        nonlocal done
        path = pdfs[index]
        docs[index] = info
        done += 1
        if not info.error:
            try:
                size, mtime = _file_signature(path)
                cache[str(path.resolve())] = (size, mtime, info)
            except OSError:
                pass
        progress("scan", done, total, f"Đã đọc {path.name}")
        if on_scan:
            on_scan(index + 1, total, path, "done", info)
        if done % CACHE_SAVE_EVERY == 0:
            _save_cache(pdf_dir, cache)

    def check_cancel() -> None:
        if CANCEL_EVENT.is_set():
            _save_cache(pdf_dir, cache)  # giữ lại phần đã quét
            raise CancelledByUser()

    def scan_sequential(indices: list[int]) -> None:
        for index in indices:
            check_cancel()
            path = pdfs[index]
            progress("scan", done, total, f"Đang đọc {path.name}")
            if on_scan:
                on_scan(index + 1, total, path, "start", None)
            try:
                info = inspect_pdf(path)
            except Exception as exc:
                info = _error_info(path, exc)
            finish(index, info)

    if pending:
        workers, threads = choose_worker_plan(len(pending))
        if workers <= 1:
            global OCR_THREADS
            OCR_THREADS = os.cpu_count() or 2
            scan_sequential(pending)
        else:
            try:
                context = mp.get_context("spawn")
                cpu_percent()  # mồi lần đo đầu (psutil trả 0.0 ở lần gọi đầu)
                queue_left = list(pending)
                in_flight = {}
                limit = min(2, workers)   # bắt đầu nhẹ nhàng, rồi tăng dần nếu còn dư
                with ProcessPoolExecutor(
                    max_workers=workers,
                    mp_context=context,
                    initializer=_init_worker,
                    initargs=(threads,),
                ) as pool:
                    while queue_left or in_flight:
                        check_cancel()
                        while queue_left and len(in_flight) < limit:
                            index = queue_left.pop(0)
                            if on_scan:
                                on_scan(index + 1, total, pdfs[index], "start", None)
                            in_flight[pool.submit(inspect_pdf, pdfs[index])] = index
                        finished, _ = wait(
                            in_flight, timeout=2, return_when=FIRST_COMPLETED
                        )
                        for future in finished:
                            index = in_flight.pop(future)
                            try:
                                info = future.result()
                            except BrokenProcessPool:
                                raise
                            except Exception as exc:
                                info = _error_info(pdfs[index], exc)
                            finish(index, info)
                        limit = adapt_concurrency(limit, workers)
            except BrokenProcessPool:
                # Hết RAM hoặc tiến trình con bị chết: quét nốt các file còn lại tuần tự.
                remaining = [i for i in pending if docs[i] is None]
                OCR_THREADS = os.cpu_count() or 2
                scan_sequential(remaining)

    _save_cache(pdf_dir, cache)
    return [doc for doc in docs if doc is not None]


def process(
    excel_path: Path,
    pdf_dir: Path,
    output_path: Path,
    progress,
    on_result=None,
    on_scan=None,
):
    df = pd.read_excel(excel_path, dtype=object)
    cols = detect_columns(df.columns)
    if not cols["name"] or not cols["id"]:
        raise ValueError("Không tự nhận ra cột Họ tên/CCCD. Hãy đổi tiêu đề cột thành 'Họ và tên' và 'Số CCCD'.")
    pdfs = sorted(pdf_dir.rglob("*.pdf"))
    if not pdfs:
        raise ValueError("Thư mục đã chọn không có file PDF.")
    docs = scan_pdfs(pdfs, pdf_dir, progress, on_scan)

    filename_matches = assign_filename_similarity_matches(
        df, cols["name"], cols["id"], docs
    )
    filename_match_owners = {
        match[0].path: row_index for row_index, match in filename_matches.items()
    }
    results, used, confidently_used = [], set(), set()
    for i, (_, row) in enumerate(df.iterrows(), 1):
        if CANCEL_EVENT.is_set():
            raise CancelledByUser()
        progress("match", i, len(df), f"Đang đối chiếu dòng {i}")
        name = str(row.get(cols["name"], "") or "")
        cid = norm_digits(row.get(cols["id"], ""))
        raw_dob = row.get(cols["dob"], "") if cols["dob"] else ""
        dob = norm_date(raw_dob)
        alternate_dob = alternate_norm_date(raw_dob)
        row_index = i - 1
        available_docs = [
            d for d in docs
            if d.path not in confidently_used
            and (
                d.path not in filename_match_owners
                or filename_match_owners[d.path] == row_index
            )
        ]
        doc, score = match_document(
            name, cid, available_docs, allow_name_fallback=False
        )
        filename_similarity_match = False
        filename_name_score = filename_id_score = 0.0
        if not doc and row_index in filename_matches:
            doc, score, filename_name_score, filename_id_score = filename_matches[row_index]
            if doc.path in confidently_used:
                doc = None
            else:
                filename_similarity_match = True
        if not doc:
            doc, score = match_document(name, cid, available_docs)
        if not doc:
            result = ["Không tìm thấy hồ sơ", "Không tìm thấy PDF khớp CCCD hoặc họ tên", "", "", ""]
            results.append(result)
            if on_result:
                on_result(i, name, cid, result)
            continue
        used.add(doc.path)
        if score >= 0.8:
            confidently_used.add(doc.path)
        mismatches = []
        format_warnings = []
        filename_id_fallback = bool(
            cid
            and any(ids_match(cid, candidate) for candidate in doc.filename_ids)
            and not (
                doc.id_from_pdf
                and any(ids_match(cid, candidate) for candidate in doc.ids)
            )
        )
        if filename_id_fallback:
            warning = "Ghép theo ID trong tên file; cần kiểm tra thủ công"
            if doc.id_from_pdf and doc.ids:
                warning = (
                    f"{warning} vì ID đọc từ PDF ({', '.join(sorted(doc.ids))}) "
                    "không khớp ID Excel"
                )
            format_warnings.append(warning)
        if filename_similarity_match:
            mismatches.append(
                "Ghép theo tên và ID gần đúng trong tên file "
                f"(độ khớp tên {filename_name_score:.0%}, "
                f"ID {filename_id_score:.0%}; ngưỡng mỗi trường 85%)"
            )
        if doc.birth_date_ambiguous:
            format_warnings.append(
                "Có nhiều ngày sinh phù hợp cùng mức tin cậy gần nhãn; cần kiểm tra thủ công"
            )
        if doc.birth_date_wide_search and doc.birth_dates:
            format_warnings.append(
                "Ngày sinh chỉ tìm thấy bằng cách tìm rộng trong cùng trang; cần kiểm tra thủ công"
            )
        if doc.birth_date_out_of_range and not doc.birth_dates:
            format_warnings.append(
                "Ngày sinh gần nhãn trong PDF ngoài độ tuổi hợp lý 18–100; cần kiểm tra thủ công"
            )
        if (
            cid
            and doc.id_from_pdf
            and doc.ids
            and not any(ids_match(cid, candidate) for candidate in doc.ids)
            and not filename_id_fallback
        ):
            mismatches.append(f"CCCD Excel {cid} không khớp PDF ({', '.join(sorted(doc.ids))})")
        if raw_dob is not None and not pd.isna(raw_dob) and str(raw_dob).strip() and not dob:
            if alternate_dob and alternate_dob in doc.birth_dates:
                format_warnings.append(
                    f"Ngày sinh Excel có thể bị đảo ngày/tháng; "
                    f"đổi thứ tự thành {alternate_dob} thì khớp với PDF"
                )
            else:
                format_warnings.append("Không nhận diện được định dạng ngày sinh trong Excel; cần kiểm tra lại ngày/tháng")
        elif dob and doc.birth_dates and dob not in doc.birth_dates:
            if alternate_dob and alternate_dob in doc.birth_dates:
                format_warnings.append(
                    f"Ngày sinh Excel có thể bị đảo ngày/tháng; "
                    f"đổi thứ tự thành {alternate_dob} thì khớp với PDF"
                )
            else:
                mismatches.append(f"Ngày sinh Excel {dob} không khớp PDF ({', '.join(sorted(doc.birth_dates))})")
        ocr_warning = ""
        if doc.used_ocr and doc.confidence is None:
            ocr_warning = "Không tính được điểm tin cậy OCR; cần kiểm tra thủ công"
        elif doc.used_ocr and doc.confidence < 50:
            uncertain = []
            if cid and not any(ids_match(cid, candidate) for candidate in doc.ids):
                uncertain.append("CCCD")
            if dob and not doc.birth_dates:
                uncertain.append("ngày sinh")
            if name and norm_text(name) not in norm_text(doc.text):
                uncertain.append("họ tên")
            if not uncertain:
                uncertain = ["CCCD, họ tên và ngày sinh"]
            ocr_warning = (
                f"Điểm OCR thấp ({doc.confidence:.1f}); thông tin chưa chắc chắn: {', '.join(uncertain)}"
            )
        if doc.error:
            status, detail = "Cần kiểm tra", "; ".join([doc.error, *format_warnings, ocr_warning])
        elif mismatches:
            if filename_similarity_match:
                status = "Sai thông tin"
            elif doc.confidence is not None and doc.confidence >= 80:
                status = "Sai thông tin"
            else:
                status = "Cần kiểm tra"
                format_warnings.append(
                    "Độ tin cậy PDF dưới 80 hoặc không có điểm; sai lệch chưa thể xác nhận"
                )
            detail = "; ".join([*mismatches, *format_warnings, ocr_warning])
        else:
            review_reasons = list(format_warnings)
            if dob and not doc.birth_dates:
                review_reasons.append("Không tìm thấy ngày sinh trong PDF để xác minh")
            if ocr_warning:
                review_reasons.append(ocr_warning)
            if score < 0.8:
                review_reasons.append(f"Độ khớp họ tên thấp ({score:.2f})")
            if review_reasons:
                status, detail = "Cần kiểm tra", "; ".join(review_reasons)
            else:
                status, detail = "OK", "Các trường có dữ liệu trong PDF đều khớp"
        confidence = round(doc.confidence, 1) if doc.confidence is not None else ""
        result = [status, detail, confidence, str(doc.path.resolve()), ""]
        results.append(result)
        if on_result:
            on_result(i, name, cid, result)

    progress("export", 1, 1, "Đang xuất file kết quả Excel...")
    for idx, col in enumerate(OUTPUT_COLUMNS):
        df[col] = [r[idx] for r in results]
    unused = [d.path for d in docs if d.path not in used]
    if unused:
        print("PDF không khớp dòng Excel:")
        for path in unused:
            print(path.name)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Kết quả đối soát")
        pd.DataFrame({"PDF không khớp dòng Excel": [str(path.resolve()) for path in unused]}).to_excel(
            writer, index=False, sheet_name="PDF chưa ghép"
        )
    style_output(output_path, cols["name"])
    counts = df["Trạng thái đối soát"].value_counts().to_dict()
    return counts, len(pdfs)


def style_output(path: Path, name_column: str):
    wb = load_workbook(path)
    ws = wb["Kết quả đối soát"]
    headers = {cell.value: cell.column for cell in ws[1]}
    fills = {
        "OK": "C6EFCE", "Cần kiểm tra": "FFF2CC",
        "Sai thông tin": "F4CCCC", "Không tìm thấy hồ sơ": "D9D9D9",
    }
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True)
    status_col = headers["Trạng thái đối soát"]
    link_col = headers["Đường dẫn tài liệu"]
    name_col = headers.get(name_column)
    for row in range(2, ws.max_row + 1):
        status = ws.cell(row, status_col).value
        ws.cell(row, status_col).fill = PatternFill("solid", fgColor=fills.get(status, "FFFFFF"))
        link = ws.cell(row, link_col).value
        if link:
            target = ws.cell(row, name_col) if name_col else ws.cell(row, link_col)
            target.hyperlink = Path(link).as_uri()
            target.font = Font(color="0563C1", underline="single")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col_cells in ws.columns:
        width = min(55, max(12, max(len(str(c.value or "")) for c in col_cells) + 2))
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = width
    wb.save(path)


# =============================================================================
# Giao diện
# =============================================================================
P = {
    "bg": "#F3F6FA",
    "card": "#FFFFFF",
    "border": "#D9E1EC",
    "primary": "#1F4E78",
    "primary_hover": "#2A6AA3",
    "text": "#1F2937",
    "muted": "#6B7280",
    "ok": "#C6EFCE", "ok_fg": "#1E6B34", "ok_row": "#EAF7EE",
    "warn": "#FFF2CC", "warn_fg": "#8A6100", "warn_row": "#FFF9E3",
    "bad": "#F4CCCC", "bad_fg": "#9B1C1C", "bad_row": "#FCE8E8",
    "none": "#E2E5E9", "none_fg": "#4B5563", "none_row": "#F0F1F3",
    "info_row": "#EAF2FB",
}
STATUS_ORDER = ["OK", "Cần kiểm tra", "Sai thông tin", "Không tìm thấy hồ sơ"]
STATUS_TAG = {
    "OK": "ok",
    "Cần kiểm tra": "warn",
    "Sai thông tin": "bad",
    "Không tìm thấy hồ sơ": "none",
}
ALL_STATUS = "Tất cả"
SETTINGS_FILE = CACHE_DIR / "settings.json"


def load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(data: dict) -> None:
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def open_path(path) -> None:
    target = str(Path(path).resolve())
    if sys.platform == "win32":
        os.startfile(target)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", target])
    else:
        subprocess.Popen(["xdg-open", target])


def fmt_time(seconds: float) -> str:
    seconds = int(max(0, seconds))
    hours, rest = divmod(seconds, 3600)
    minutes, secs = divmod(rest, 60)
    return f"{hours}:{minutes:02d}:{secs:02d}" if hours else f"{minutes:02d}:{secs:02d}"


class App(tk.Tk):
    STEPS = [("scan", "1  Quét PDF"), ("match", "2  Đối chiếu"), ("export", "3  Xuất Excel")]

    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.configure(bg=P["bg"])
        self.minsize(980, 680)
        width, height = 1160, 820
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        y = max(0, (self.winfo_screenheight() - height) // 3)
        self.geometry(f"{width}x{height}+{x}+{y}")

        saved = load_settings()
        self.excel = tk.StringVar(value=saved.get("excel", ""))
        self.folder = tk.StringVar(value=saved.get("folder", ""))
        self.output = tk.StringVar(value=saved.get("output", ""))
        self.use_cache = tk.BooleanVar(value=saved.get("cache", True))
        self.status_filter = tk.StringVar(value=ALL_STATUS)
        self.search_match = tk.StringVar()
        self.search_scan = tk.StringVar()
        self.status_var = tk.StringVar(value="Sẵn sàng. Chọn file Excel và thư mục PDF rồi bấm Bắt đầu.")
        self.timing_var = tk.StringVar(value="")
        self.scan_stats_var = tk.StringVar(value="")
        self.match_count_var = tk.StringVar(value="")
        self.scan_count_var = tk.StringVar(value="")

        self.queue: Queue = Queue()
        self.running = False
        self.scan_rows: dict[str, tuple[tuple, str]] = {}
        self.scan_paths: dict[str, str] = {}
        self.match_rows: dict[str, tuple[tuple, str]] = {}
        self.counts: Counter = Counter()
        self.scan_stats: Counter = Counter()
        self.phase: str | None = None
        self.phase_t0: dict[str, float] = {}
        self.run_t0 = 0.0
        self.progress_n = 0
        self.progress_total = 0
        self.fraction = 0.0
        self._timers: dict[str, str] = {}
        self._titles: dict[tuple[int, str], str] = {}
        self._sort_dir: dict[tuple[int, str], bool] = {}
        self.input_widgets: list = []
        self.tiles: dict[str, tk.Label] = {}

        self._setup_style()
        self._build()
        for var in (self.excel, self.folder):
            var.trace_add("write", lambda *_: self._debounce("info", self._refresh_info, 500))
        self.search_match.trace_add("write", lambda *_: self._debounce("fm", self._rebuild_match))
        self.search_scan.trace_add("write", lambda *_: self._debounce("fs", self._rebuild_scan))
        self._refresh_info()
        self.bind_all("<Button-1>", self._on_global_click, add="+")
        self.bind_class("TEntry", "<FocusOut>", lambda e: e.widget.selection_clear(), add="+")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ style
    def _setup_style(self):
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
            try:
                tkfont.nametofont(name).configure(family="Segoe UI", size=10)
            except tk.TclError:
                pass
        style = ttk.Style(self)
        style.theme_use("clam")
        ui = "Segoe UI"

        def fixed_map(option):
            # Sửa lỗi Tk 8.6.9: tag màu của Treeview không hiển thị.
            return [
                entry for entry in style.map("Treeview", query_opt=option)
                if entry[:2] != ("!disabled", "!selected")
            ]

        style.configure(".", background=P["bg"], foreground=P["text"], font=(ui, 10))
        style.configure("TFrame", background=P["bg"])
        style.configure("Card.TFrame", background=P["card"])
        style.configure("TLabel", background=P["bg"])
        style.configure("Card.TLabel", background=P["card"])
        style.configure("Muted.TLabel", background=P["card"], foreground=P["muted"], font=(ui, 9))
        style.configure("Good.TLabel", background=P["card"], foreground=P["ok_fg"], font=(ui, 9))
        style.configure("Bad.TLabel", background=P["card"], foreground=P["bad_fg"], font=(ui, 9))
        style.configure("Bar.TLabel", background=P["bg"], foreground=P["muted"], font=(ui, 9))
        style.configure("TCheckbutton", background=P["bg"], focuscolor=P["bg"])
        style.map("TCheckbutton", background=[("active", P["bg"])])
        style.configure(
            "TEntry", padding=7, fieldbackground="#FFFFFF", foreground=P["text"],
            insertcolor="#000000", insertwidth=2,
            selectbackground="#BBD6F0", selectforeground=P["text"],
            bordercolor=P["border"], lightcolor=P["border"], darkcolor=P["border"],
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", P["primary"])],
            lightcolor=[("focus", P["primary"])],
            darkcolor=[("focus", P["primary"])],
        )
        style.configure(
            "TButton", padding=(14, 7), background="#FFFFFF", foreground=P["text"],
            bordercolor=P["border"], lightcolor="#FFFFFF", darkcolor="#FFFFFF",
            focuscolor="#FFFFFF",
        )
        style.map(
            "TButton",
            background=[("disabled", "#F3F4F6"), ("active", "#EAF1F8")],
            foreground=[("disabled", "#9CA3AF")],
        )
        style.configure(
            "Primary.TButton", font=(ui, 11, "bold"), padding=(28, 10),
            background=P["primary"], foreground="#FFFFFF", bordercolor=P["primary"],
            lightcolor=P["primary"], darkcolor=P["primary"], focuscolor=P["primary"],
        )
        style.map(
            "Primary.TButton",
            background=[("disabled", "#9DB3C8"), ("active", P["primary_hover"])],
            foreground=[("disabled", "#FFFFFF")],
        )
        style.configure(
            "Danger.TButton", font=(ui, 10, "bold"), padding=(18, 10),
            background="#FFFFFF", foreground="#B42318", bordercolor="#F0B4AE",
        )
        style.map(
            "Danger.TButton",
            background=[("disabled", "#F3F4F6"), ("active", "#FDECEA")],
            foreground=[("disabled", "#9CA3AF")],
        )
        style.configure(
            "Accent.Horizontal.TProgressbar", troughcolor="#E5EAF1",
            background=P["primary"], bordercolor="#E5EAF1",
            lightcolor=P["primary"], darkcolor=P["primary"], thickness=16,
        )
        style.configure(
            "Treeview", rowheight=27, background="#FFFFFF", fieldbackground="#FFFFFF",
            borderwidth=0, font=(ui, 10),
        )
        style.configure(
            "Treeview.Heading", font=(ui, 10, "bold"), background="#EEF2F7",
            foreground=P["text"], padding=(8, 7), relief="flat",
        )
        style.map("Treeview.Heading", background=[("active", "#E1E8F0")])
        style.map(
            "Treeview",
            foreground=fixed_map("foreground"),
            background=[("selected", "focus", "#BBD6F0")] + fixed_map("background"),
        )
        style.configure("TNotebook", background=P["bg"], borderwidth=0)
        style.configure(
            "TNotebook.Tab", padding=(20, 9), font=(ui, 10, "bold"),
            background="#E5EAF1", foreground=P["muted"],
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#FFFFFF")],
            foreground=[("selected", P["primary"])],
            focuscolor=[("selected", "#FFFFFF"), ("!selected", "#E5EAF1")],
        )
        style.configure("TCombobox", padding=5, insertcolor="#000000", insertwidth=2)
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#FFFFFF")],
            selectbackground=[("readonly", "!focus", "#FFFFFF")],
            selectforeground=[("readonly", "!focus", P["text"])],
        )
        style.configure("Vertical.TScrollbar", background="#E5EAF1", troughcolor=P["card"], bordercolor=P["card"])
        style.configure("Horizontal.TScrollbar", background="#E5EAF1", troughcolor=P["card"], bordercolor=P["card"])

    # ----------------------------------------------------------------- layout
    def _card(self, parent) -> tk.Frame:
        return tk.Frame(parent, bg=P["card"], highlightbackground=P["border"], highlightthickness=1)

    def _build(self):
        header = tk.Frame(self, bg=P["primary"])
        header.pack(fill="x")
        left = tk.Frame(header, bg=P["primary"])
        left.pack(side="left", padx=26, pady=14)
        tk.Label(left, text="ĐỐI SOÁT HỒ SƠ", bg=P["primary"], fg="#FFFFFF",
                 font=("Segoe UI", 19, "bold")).pack(anchor="w")
        tk.Label(left, text="Tự động quét PDF, OCR tiếng Việt và đối chiếu với danh sách Excel",
                 bg=P["primary"], fg="#C9DAEA", font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))
        tk.Label(header, text="Quét song song  •  Có cache  •  Hủy bất cứ lúc nào",
                 bg=P["primary"], fg="#9FB9D3", font=("Segoe UI", 9)).pack(side="right", padx=26)

        body = ttk.Frame(self, padding=(22, 16, 22, 16))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(4, weight=1)

        self._build_inputs(body)
        self._build_actions(body)
        self._build_progress(body)
        self._build_tiles(body)
        self._build_results(body)

    def _build_inputs(self, body):
        card = self._card(body)
        card.grid(row=0, column=0, sticky="ew")
        inner = ttk.Frame(card, style="Card.TFrame", padding=(18, 12, 18, 12))
        inner.pack(fill="x")
        inner.columnconfigure(1, weight=1)
        inner.columnconfigure(0, minsize=150)
        specs = [
            ("File Excel (danh sách)", self.excel, self.pick_excel, "excel_info"),
            ("Thư mục PDF", self.folder, self.pick_folder, "folder_info"),
            ("File kết quả", self.output, self.pick_output, None),
        ]
        for index, (label, variable, command, info_name) in enumerate(specs):
            row = index * 2
            ttk.Label(inner, text=label, style="Card.TLabel",
                      font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", pady=(6, 0))
            entry = ttk.Entry(inner, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", pady=(6, 0))
            button = ttk.Button(inner, text="Chọn...", command=command)
            button.grid(row=row, column=2, padx=(10, 0), pady=(6, 0))
            self.input_widgets += [entry, button]
            if info_name:
                label_widget = ttk.Label(inner, text="", style="Muted.TLabel")
                label_widget.grid(row=row + 1, column=1, sticky="w", pady=(2, 0))
                setattr(self, info_name, label_widget)

    def _build_actions(self, body):
        bar = ttk.Frame(body)
        bar.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        self.start = ttk.Button(bar, text="Bắt đầu kiểm tra", style="Primary.TButton", command=self.run)
        self.start.pack(side="left")
        self.cancel_btn = ttk.Button(bar, text="Hủy", style="Danger.TButton",
                                     command=self.cancel, state="disabled")
        self.cancel_btn.pack(side="left", padx=(10, 0))
        self.cache_check = ttk.Checkbutton(bar, text="Dùng cache quét PDF", variable=self.use_cache)
        self.cache_check.pack(side="left", padx=(26, 0))
        self.clear_btn = ttk.Button(bar, text="Xóa cache", command=self.clear_cache)
        self.clear_btn.pack(side="left", padx=(12, 0))
        self.open_dir_btn = ttk.Button(bar, text="Mở thư mục", command=self.open_result_folder, state="disabled")
        self.open_dir_btn.pack(side="right")
        self.open_file_btn = ttk.Button(bar, text="Mở file kết quả", command=self.open_result_file, state="disabled")
        self.open_file_btn.pack(side="right", padx=(0, 8))

    def _build_progress(self, body):
        card = self._card(body)
        card.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        inner = ttk.Frame(card, style="Card.TFrame", padding=(18, 12, 18, 12))
        inner.pack(fill="x")
        inner.columnconfigure(0, weight=1)

        steps = ttk.Frame(inner, style="Card.TFrame")
        steps.grid(row=0, column=0, columnspan=2, sticky="w")
        self.step_labels: dict[str, tk.Label] = {}
        for index, (key, text) in enumerate(self.STEPS):
            label = tk.Label(steps, text=text, bg=P["card"], fg=P["muted"], font=("Segoe UI", 10))
            label.pack(side="left")
            self.step_labels[key] = label
            if index < len(self.STEPS) - 1:
                tk.Label(steps, text="   ›   ", bg=P["card"], fg="#B8C2CF").pack(side="left")

        self.bar = ttk.Progressbar(inner, style="Accent.Horizontal.TProgressbar",
                                   mode="determinate", maximum=1000)
        self.bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 8))
        ttk.Label(inner, textvariable=self.status_var, style="Card.TLabel",
                  wraplength=720).grid(row=2, column=0, sticky="w")
        ttk.Label(inner, textvariable=self.timing_var, style="Card.TLabel",
                  font=("Segoe UI", 10, "bold")).grid(row=2, column=1, sticky="e")
        ttk.Label(inner, textvariable=self.scan_stats_var, style="Muted.TLabel").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

    def _build_tiles(self, body):
        row = ttk.Frame(body)
        row.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        specs = [
            ("total", "Tổng số dòng", "#FFFFFF", P["text"]),
            ("OK", "OK", P["ok"], P["ok_fg"]),
            ("Cần kiểm tra", "Cần kiểm tra", P["warn"], P["warn_fg"]),
            ("Sai thông tin", "Sai thông tin", P["bad"], P["bad_fg"]),
            ("Không tìm thấy hồ sơ", "Không tìm thấy hồ sơ", P["none"], P["none_fg"]),
        ]
        for column, (key, title, background, foreground) in enumerate(specs):
            row.columnconfigure(column, weight=1, uniform="tile")
            tile = tk.Frame(row, bg=background, highlightbackground=P["border"],
                            highlightthickness=1, cursor="hand2")
            tile.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
            number = tk.Label(tile, text="0", bg=background, fg=foreground,
                              font=("Segoe UI", 22, "bold"), cursor="hand2")
            number.pack(anchor="w", padx=14, pady=(8, 0))
            caption = tk.Label(tile, text=title, bg=background, fg=foreground,
                               font=("Segoe UI", 10), cursor="hand2")
            caption.pack(anchor="w", padx=14, pady=(0, 8))
            self.tiles[key] = number
            for widget in (tile, number, caption):
                widget.bind("<Button-1>", lambda _e, k=key: self._filter_by(k))

    def _build_results(self, body):
        wrap = ttk.Frame(body)
        wrap.grid(row=4, column=0, sticky="nsew", pady=(14, 0))
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.notebook = ttk.Notebook(wrap)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        match_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=10)
        scan_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=10)
        self.notebook.add(match_tab, text="Kết quả đối soát")
        self.notebook.add(scan_tab, text="Quét PDF")

        # --- tab đối soát
        toolbar = ttk.Frame(match_tab, style="Card.TFrame")
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(toolbar, text="Tìm:", style="Card.TLabel").pack(side="left")
        ttk.Entry(toolbar, textvariable=self.search_match, width=30).pack(side="left", padx=(6, 14))
        ttk.Label(toolbar, text="Trạng thái:", style="Card.TLabel").pack(side="left")
        combo = ttk.Combobox(toolbar, textvariable=self.status_filter, state="readonly", width=22,
                             values=[ALL_STATUS, *STATUS_ORDER])
        combo.pack(side="left", padx=(6, 14))
        combo.bind("<<ComboboxSelected>>", lambda _e: (self._rebuild_match(), combo.selection_clear(), self.focus_set()))
        ttk.Label(toolbar, textvariable=self.match_count_var, style="Muted.TLabel").pack(side="left")
        ttk.Label(toolbar, text="Nhấp đúp để mở PDF  •  Nhấp phải để thêm tùy chọn  •  Nhấp tiêu đề cột để sắp xếp",
                  style="Muted.TLabel").pack(side="right")
        self.match_tree = self._make_tree(
            match_tab,
            {
                "row": ("Dòng", 55, "e", False),
                "name": ("Họ tên", 190, "w", False),
                "cid": ("CCCD", 120, "w", False),
                "status": ("Trạng thái", 135, "w", False),
                "conf": ("OCR (%)", 70, "e", False),
                "detail": ("Chi tiết", 360, "w", True),
                "path": ("Tài liệu", 260, "w", True),
            },
        )

        # --- tab quét PDF
        toolbar = ttk.Frame(scan_tab, style="Card.TFrame")
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(toolbar, text="Tìm:", style="Card.TLabel").pack(side="left")
        ttk.Entry(toolbar, textvariable=self.search_scan, width=30).pack(side="left", padx=(6, 14))
        ttk.Label(toolbar, textvariable=self.scan_count_var, style="Muted.TLabel").pack(side="left")
        self.scan_tree = self._make_tree(
            scan_tab,
            {
                "scan": ("Lượt", 55, "e", False),
                "pdf": ("PDF", 250, "w", False),
                "status": ("Trạng thái quét", 125, "w", False),
                "ids": ("CCCD/CMND", 130, "w", False),
                "id_score": ("Điểm ID", 70, "e", False),
                "dob": ("Ngày sinh", 100, "w", False),
                "dob_score": ("Điểm NS", 70, "e", False),
                "conf": ("OCR (%)", 70, "e", False),
                "text": ("Văn bản trích xuất", 400, "w", True),
            },
        )

    def _make_tree(self, parent, spec: dict) -> ttk.Treeview:
        holder = ttk.Frame(parent, style="Card.TFrame")
        holder.pack(fill="both", expand=True)
        tree = ttk.Treeview(holder, columns=tuple(spec), show="headings", selectmode="browse")
        for column, (title, width, anchor, stretch) in spec.items():
            self._titles[(id(tree), column)] = title
            tree.heading(column, text=title, command=lambda c=column, t=tree: self._sort(t, c))
            tree.column(column, width=width, minwidth=50, anchor=anchor, stretch=stretch)
        vertical = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(holder, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        for tag in ("ok", "warn", "bad", "none"):
            tree.tag_configure(tag, background=P[f"{tag}_row"])
        tree.tag_configure("ocr", background=P["info_row"])
        tree.tag_configure("err", background=P["bad_row"], foreground=P["bad_fg"])
        tree.bind(
            "<Button-1>",
            lambda e, t=tree: t.selection_remove(*t.selection()) if not t.identify_row(e.y) else None,
            add="+",
        )
        tree.bind("<Double-1>", lambda _e, t=tree: self._open_selected(t))
        tree.bind("<Button-3>", lambda e, t=tree: self._context_menu(e, t))
        tree.bind("<Button-2>", lambda e, t=tree: self._context_menu(e, t))
        return tree

    # ------------------------------------------------------------ small utils
    def _reset_widget_states(self):
        """Bỏ trạng thái hover/nhấn còn kẹt của nút và đưa focus ra khỏi ô nhập."""
        def walk(widget):
            for child in widget.winfo_children():
                if isinstance(child, (ttk.Button, ttk.Checkbutton)):
                    child.state(["!active", "!pressed"])
                walk(child)

        walk(self)
        self.focus_set()

    def _on_global_click(self, event):
        widget = event.widget
        if not isinstance(widget, tk.Misc):
            return  # ví dụ: danh sách thả xuống của combobox
        if not isinstance(widget, (ttk.Treeview, ttk.Scrollbar, tk.Menu)):
            # Bấm ra ngoài bảng (kể cả vào nút, ô tìm kiếm...) -> bỏ chọn dòng.
            for tree in (self.match_tree, self.scan_tree):
                selected = tree.selection()
                if selected:
                    tree.selection_remove(*selected)
        interactive = (ttk.Entry, ttk.Combobox, ttk.Treeview, ttk.Button, ttk.Checkbutton,
                       ttk.Scrollbar, ttk.Notebook, tk.Menu, tk.Listbox, tk.Entry)
        if isinstance(widget, interactive):
            return
        self._reset_widget_states()  # bấm vào chỗ trống: mọi thứ về trạng thái bình thường

    def _debounce(self, key: str, func, delay: int = 250):
        old = self._timers.get(key)
        if old:
            self.after_cancel(old)
        self._timers[key] = self.after(delay, func)

    def _save_settings(self):
        save_settings({
            "excel": self.excel.get(),
            "folder": self.folder.get(),
            "output": self.output.get(),
            "cache": self.use_cache.get(),
        })

    # ----------------------------------------------------------------- inputs
    def _refresh_info(self):
        excel = self.excel.get().strip()
        if not excel:
            self.excel_info.config(text="", style="Muted.TLabel")
        elif not Path(excel).is_file():
            self.excel_info.config(text="Không tìm thấy file Excel này", style="Bad.TLabel")
        else:
            try:
                header = pd.read_excel(excel, nrows=0, dtype=object)
                cols = detect_columns(header.columns)
                if cols["name"] and cols["id"]:
                    text = f"Nhận diện cột:  Họ tên = “{cols['name']}”   •   CCCD = “{cols['id']}”"
                    text += f"   •   Ngày sinh = “{cols['dob']}”" if cols["dob"] else "   •   (không có cột ngày sinh)"
                    self.excel_info.config(text=text, style="Good.TLabel")
                else:
                    self.excel_info.config(
                        text="Không nhận ra cột Họ tên/CCCD. Hãy đổi tiêu đề thành “Họ và tên” và “Số CCCD”.",
                        style="Bad.TLabel",
                    )
            except Exception as exc:
                self.excel_info.config(text=f"Không đọc được file Excel: {exc}", style="Bad.TLabel")

        folder = self.folder.get().strip()
        if not folder:
            self.folder_info.config(text="", style="Muted.TLabel")
        elif not Path(folder).is_dir():
            self.folder_info.config(text="Không tìm thấy thư mục này", style="Bad.TLabel")
        else:
            self.folder_info.config(text="Đang đếm file PDF...", style="Muted.TLabel")

            def count():
                try:
                    total = sum(1 for _ in Path(folder).rglob("*.pdf"))
                except Exception:
                    total = -1
                try:
                    self.after(0, lambda: self._show_pdf_count(folder, total))
                except Exception:
                    pass

            threading.Thread(target=count, daemon=True).start()

    def _show_pdf_count(self, folder: str, total: int):
        if folder != self.folder.get().strip():
            return
        if total < 0:
            self.folder_info.config(text="Không đọc được thư mục", style="Bad.TLabel")
        elif total == 0:
            self.folder_info.config(text="Thư mục không có file PDF nào", style="Bad.TLabel")
        else:
            self.folder_info.config(text=f"Tìm thấy {total:,} file PDF (gồm cả thư mục con)", style="Good.TLabel")

    def pick_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.excel.set(path)
            self.output.set(str(Path(path).with_name(Path(path).stem + "_ket_qua.xlsx")))
        self._reset_widget_states()

    def pick_folder(self):
        if path := filedialog.askdirectory():
            self.folder.set(path)
        self._reset_widget_states()

    def pick_output(self):
        if path := filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]):
            self.output.set(path)
        self._reset_widget_states()

    # ---------------------------------------------------------------- running
    def _set_running(self, running: bool):
        self.running = running
        state = "disabled" if running else "normal"
        self.start.config(state=state)
        self.clear_btn.config(state=state)
        self.cache_check.config(state=state)  # chỉ đọc một lần lúc bấm Bắt đầu
        for widget in self.input_widgets:
            widget.config(state=state)
        self.cancel_btn.config(state="normal" if running else "disabled", text="Hủy")
        if running:
            self.open_file_btn.config(state="disabled")
            self.open_dir_btn.config(state="disabled")

    def _reset_results(self):
        for tree in (self.match_tree, self.scan_tree):
            tree.delete(*tree.get_children())
        self.scan_rows.clear()
        self.scan_paths.clear()
        self.match_rows.clear()
        self.counts.clear()
        self.scan_stats.clear()
        self.phase = None
        self.phase_t0.clear()
        self.progress_n = self.progress_total = 0
        self.fraction = 0.0
        self.bar.config(value=0)
        self.status_filter.set(ALL_STATUS)
        self.search_match.set("")
        self.search_scan.set("")
        self.scan_stats_var.set("")
        self.timing_var.set("")
        self.match_count_var.set("")
        self.scan_count_var.set("")
        self._update_tiles()
        self._set_steps(None)

    def run(self):
        if self.running:
            return
        excel, folder, output = self.excel.get().strip(), self.folder.get().strip(), self.output.get().strip()
        if not excel or not folder or not output:
            messagebox.showwarning(APP_NAME, "Vui lòng chọn đủ file Excel, thư mục PDF và file kết quả.")
            return
        if not Path(excel).is_file():
            messagebox.showerror(APP_NAME, "Không tìm thấy file Excel đã chọn.")
            return
        if not Path(folder).is_dir():
            messagebox.showerror(APP_NAME, "Không tìm thấy thư mục PDF đã chọn.")
            return
        if Path(output).resolve() == Path(excel).resolve():
            messagebox.showerror(APP_NAME, "File kết quả không được trùng với file Excel nguồn.")
            return
        confirmed = confirm_output_path(Path(output))
        if confirmed is None:
            return
        self.output.set(str(confirmed))
        self._save_settings()

        global USE_CACHE
        USE_CACHE = self.use_cache.get()
        CANCEL_EVENT.clear()
        self._reset_results()
        self._set_running(True)
        self.run_t0 = time.time()
        self.status_var.set("Đang chuẩn bị...")
        self.notebook.select(1)  # xem tiến độ quét PDF trước
        threading.Thread(
            target=self._worker, args=(Path(excel), Path(folder), confirmed), daemon=True
        ).start()
        self.after(100, self._poll)
        self.after(1000, self._tick)

    def cancel(self):
        if not self.running:
            return
        CANCEL_EVENT.set()
        self.cancel_btn.config(state="disabled", text="Đang hủy...")
        self.status_var.set("Đang hủy: chờ các PDF đang xử lý chạy xong (kết quả đã quét vẫn được lưu cache)...")

    def _worker(self, excel: Path, folder: Path, output: Path):
        put = self.queue.put
        try:
            result = process(
                excel, folder, output,
                lambda phase, n, total, msg: put(("progress", phase, n, total, msg)),
                lambda row, name, cid, values: put(("result", row, name, cid, *values)),
                lambda n, total, path, phase, info: put(("scan", n, total, path, phase, info)),
            )
            put(("done", result))
        except CancelledByUser:
            put(("cancelled",))
        except Exception as exc:
            error_log = CACHE_DIR / "errors.log"
            try:
                CACHE_DIR.mkdir(parents=True, exist_ok=True)
                with error_log.open("a", encoding="utf-8") as log_file:
                    log_file.write(
                        f"\n{'=' * 80}\n{datetime.now().isoformat()}\n"
                        f"{traceback.format_exc()}"
                    )
            except OSError as log_error:
                put((
                    "error",
                    f"{exc}\n\nKhông lưu được traceback: {log_error}",
                ))
            else:
                put((
                    "error",
                    f"{exc}\n\nTraceback đầy đủ đã lưu tại:\n{error_log}",
                ))

    # ---------------------------------------------------------------- polling
    def _poll(self):
        handled = 0
        try:
            while handled < 300:
                event = self.queue.get_nowait()
                handled += 1
                if self._handle(event) == "stop":
                    return
        except Empty:
            pass
        self.after(30 if handled >= 300 else 100, self._poll)

    def _tick(self):
        if not self.running:
            return
        self._update_timing()
        self.notebook.tab(1, text=f"Quét PDF ({len(self.scan_rows)})")
        self.after(1000, self._tick)

    def _handle(self, event):
        kind = event[0]
        if kind == "progress":
            _, phase, n, total, message = event
            self._update_progress(phase, n, total, message)
        elif kind == "scan":
            self._on_scan(event)
        elif kind == "result":
            self._on_result(event)
        elif kind == "done":
            self._finish(event[1])
            return "stop"
        elif kind == "cancelled":
            self._set_running(False)
            self.status_var.set("Đã hủy. Phần PDF đã quét được lưu cache, lần chạy sau sẽ nhanh hơn.")
            self.timing_var.set("")
            return "stop"
        elif kind == "error":
            self._set_running(False)
            self.status_var.set("Có lỗi xảy ra.")
            messagebox.showerror(APP_NAME, event[1])
            return "stop"
        return None

    def _update_progress(self, phase: str, n: int, total: int, message: str):
        now = time.time()
        if phase != self.phase:
            self.phase = phase
            self.phase_t0[phase] = now
            self._set_steps(phase)
            if phase == "match" and self.notebook.index("current") == 1:
                self.notebook.select(0)  # chuyển sang bảng đối soát khi bắt đầu đối chiếu
        self.progress_n, self.progress_total = n, total
        ratio = n / total if total else 1.0
        if phase == "scan":
            self.fraction = 0.85 * ratio
        elif phase == "match":
            self.fraction = 0.85 + 0.13 * ratio
        else:
            self.fraction = 1.0
        self.bar.config(value=self.fraction * 1000)
        self.status_var.set(message)
        self._update_timing()

    def _update_timing(self):
        elapsed = time.time() - self.run_t0
        n, total = self.progress_n, self.progress_total
        text = f"{self.fraction * 100:.0f}%   •   Đã chạy {fmt_time(elapsed)}"
        if self.phase == "scan" and 2 <= n < total:
            per_item = (time.time() - self.phase_t0.get("scan", self.run_t0)) / n
            text += f"   •   Còn ~{fmt_time(per_item * (total - n))}"
        self.timing_var.set(text)

    def _set_steps(self, phase: str | None, all_done: bool = False):
        keys = [key for key, _ in self.STEPS]
        current = keys.index(phase) if phase in keys else -1
        for index, (key, text) in enumerate(self.STEPS):
            label = self.step_labels[key]
            if all_done or index < current:
                label.config(text="✔ " + text, fg=P["ok_fg"], font=("Segoe UI", 10))
            elif index == current:
                label.config(text=text, fg=P["primary"], font=("Segoe UI", 10, "bold"))
            else:
                label.config(text=text, fg=P["muted"], font=("Segoe UI", 10))

    # ----------------------------------------------------------- scan / match
    def _upsert(self, tree, item_id: str, values: tuple, tag: str):
        if tree.exists(item_id):
            tree.item(item_id, values=values, tags=(tag,) if tag else ())
            return
        at_bottom = tree.yview()[1] >= 0.97
        tree.insert("", "end", iid=item_id, values=values, tags=(tag,) if tag else ())
        if at_bottom:
            tree.see(item_id)

    def _on_scan(self, event):
        _, index, _total, path, phase, info = event
        item_id = f"scan-{index}"
        self.scan_paths[item_id] = str(path)
        if phase == "start":
            values = (index, path.name, "Đang quét...", "", "", "", "", "", "")
            tag = ""
        else:
            if info.error:
                status, tag = "Lỗi đọc", "err"
                self.scan_stats["err"] += 1
            elif info.used_ocr:
                status, tag = "Đã OCR", "ocr"
                self.scan_stats["ocr"] += 1
            else:
                status, tag = "Đã đọc text", ""
                self.scan_stats["text"] += 1
            confidence = round(info.confidence, 1) if info.confidence is not None else ""
            values = (
                index,
                path.name,
                status,
                ", ".join(sorted(info.ids)),
                f"{info.id_confidence:.1f}" if info.id_confidence is not None else "",
                ", ".join(sorted(info.birth_dates)),
                f"{info.birth_date_confidence:.1f}" if info.birth_date_confidence is not None else "",
                confidence,
                " ".join(info.text[:300].split()),
            )
            self.scan_stats_var.set(
                f"Đã quét:  {self.scan_stats['ocr']} OCR   •   {self.scan_stats['text']} đọc text   "
                f"•   {self.scan_stats['err']} lỗi"
            )
        self.scan_rows[f"scan-{index}"] = (values, tag)
        if self._scan_visible(values):
            self._upsert(self.scan_tree, item_id, values, tag)

    def _on_result(self, event):
        _, row, name, cid, status, detail, confidence, path, _note = event
        item_id = f"row-{row}"
        values = (row, name, cid, status, confidence, detail, path)
        tag = STATUS_TAG.get(status, "")
        self.match_rows[item_id] = (values, tag)
        self.counts[status] += 1
        self._update_tiles()
        if self._match_visible(values):
            self._upsert(self.match_tree, item_id, values, tag)
            self._update_match_counter(len(self.match_tree.get_children()))

    def _update_tiles(self):
        self.tiles["total"].config(text=f"{sum(self.counts.values()):,}")
        for status in STATUS_ORDER:
            self.tiles[status].config(text=f"{self.counts.get(status, 0):,}")

    def _finish(self, result):
        counts, pdf_count = result
        self.counts = Counter(counts)
        self._update_tiles()
        self._set_running(False)
        self.fraction = 1.0
        self.bar.config(value=1000)
        self._set_steps(None, all_done=True)
        self._rebuild_match()
        self._rebuild_scan()
        self.notebook.tab(1, text=f"Quét PDF ({len(self.scan_rows)})")
        self.notebook.select(0)
        elapsed = time.time() - self.run_t0
        self.status_var.set(f"Hoàn tất {pdf_count:,} PDF, {sum(counts.values()):,} dòng Excel. File đã lưu: {self.output.get()}")
        self.timing_var.set(f"100%   •   Tổng thời gian {fmt_time(elapsed)}")
        self.open_file_btn.config(state="normal")
        self.open_dir_btn.config(state="normal")
        self.bell()

    # ------------------------------------------------------- filter and sort
    def _match_visible(self, values: tuple) -> bool:
        wanted = self.status_filter.get()
        if wanted != ALL_STATUS and values[3] != wanted:
            return False
        query = norm_text(self.search_match.get())
        if not query:
            return True
        haystack = norm_text(" ".join(str(values[i]) for i in (1, 2, 5, 6)))
        return query in haystack

    def _scan_visible(self, values: tuple) -> bool:
        query = norm_text(self.search_scan.get())
        if not query:
            return True
        return query in norm_text(" ".join(str(values[i]) for i in (1, 3, 5)))

    def _update_match_counter(self, shown: int):
        filtering = bool(self.search_match.get().strip()) or self.status_filter.get() != ALL_STATUS
        self.match_count_var.set(
            f"Đang lọc: hiển thị {shown}/{len(self.match_rows)} dòng" if filtering else ""
        )

    def _update_scan_counter(self, shown: int):
        filtering = bool(self.search_scan.get().strip())
        self.scan_count_var.set(
            f"Đang lọc: hiển thị {shown}/{len(self.scan_rows)} PDF" if filtering else ""
        )

    def _rebuild_match(self):
        tree = self.match_tree
        tree.delete(*tree.get_children())
        shown = 0
        for item_id, (values, tag) in sorted(self.match_rows.items(), key=lambda kv: kv[1][0][0]):
            if self._match_visible(values):
                tree.insert("", "end", iid=item_id, values=values, tags=(tag,) if tag else ())
                shown += 1
        self._update_match_counter(shown)

    def _rebuild_scan(self):
        tree = self.scan_tree
        tree.delete(*tree.get_children())
        shown = 0
        for item_id, (values, tag) in sorted(self.scan_rows.items(), key=lambda kv: kv[1][0][0]):
            if self._scan_visible(values):
                tree.insert("", "end", iid=item_id, values=values, tags=(tag,) if tag else ())
                shown += 1
        self._update_scan_counter(shown)

    def _filter_by(self, key: str):
        self.status_filter.set(ALL_STATUS if key == "total" else key)
        self.notebook.select(0)
        self._rebuild_match()

    def _sort(self, tree, column: str):
        key = (id(tree), column)
        descending = not self._sort_dir.get(key, True)
        self._sort_dir = {key: descending}

        def sort_key(item_id):
            raw = str(tree.set(item_id, column)).replace(",", "").strip()
            try:
                return (0, float(raw), "")
            except ValueError:
                return (1, 0.0, norm_text(raw))

        items = sorted(tree.get_children(""), key=sort_key, reverse=descending)
        for position, item_id in enumerate(items):
            tree.move(item_id, "", position)
        for (tree_id, col), title in self._titles.items():
            if tree_id == id(tree):
                arrow = (" ▼" if descending else " ▲") if col == column else ""
                tree.heading(col, text=title + arrow)

    # ---------------------------------------------- open / copy / context menu
    def _path_for(self, tree, item_id: str) -> str:
        if tree is self.match_tree:
            return str(tree.set(item_id, "path"))
        return self.scan_paths.get(item_id, "")

    def _open_selected(self, tree):
        selection = tree.selection()
        if not selection:
            return
        path = self._path_for(tree, selection[0])
        if not path:
            messagebox.showinfo(APP_NAME, "Dòng này chưa có tài liệu PDF để mở.")
            return
        try:
            open_path(path)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Không mở được file:\n{exc}")

    def _context_menu(self, event, tree):
        item_id = tree.identify_row(event.y)
        if not item_id:
            return
        tree.selection_set(item_id)
        path = self._path_for(tree, item_id)
        menu = tk.Menu(self, tearoff=0)
        state = "normal" if path else "disabled"
        menu.add_command(label="Mở PDF", state=state, command=lambda: self._open_selected(tree))
        menu.add_command(
            label="Mở thư mục chứa PDF", state=state,
            command=lambda: open_path(Path(path).parent),
        )
        menu.add_separator()
        menu.add_command(label="Sao chép đường dẫn", state=state, command=lambda: self._copy(path))
        cid_column = "cid" if tree is self.match_tree else "ids"
        cid_value = str(tree.set(item_id, cid_column)).strip()
        menu.add_command(
            label="Sao chép CCCD",
            state="normal" if cid_value else "disabled",  # PDF chưa quét xong thì chưa có CCCD
            command=lambda: self._copy(cid_value),
        )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _copy(self, text: str):
        self.clipboard_clear()
        self.clipboard_append(text)

    def open_result_file(self):
        try:
            open_path(self.output.get())
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Không mở được file:\n{exc}")

    def open_result_folder(self):
        try:
            open_folder(Path(self.output.get()).parent)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Không mở được thư mục:\n{exc}")

    # ------------------------------------------------------------------ cache
    def clear_cache(self):
        files = list(CACHE_DIR.glob("scan_*.pkl")) if CACHE_DIR.exists() else []
        if not files:
            messagebox.showinfo(APP_NAME, "Cache đang trống, không có gì để xóa.")
            return
        size_mb = sum(f.stat().st_size for f in files) / 2**20
        if not messagebox.askyesno(
            APP_NAME,
            f"Xóa {len(files)} file cache ({size_mb:.1f} MB)?\n"
            "Lần quét tiếp theo sẽ phải OCR lại toàn bộ PDF.",
        ):
            return
        failed = 0
        for file in files:
            try:
                file.unlink()
            except OSError:
                failed += 1
        messagebox.showinfo(
            APP_NAME,
            "Đã xóa cache." if not failed else f"Đã xóa cache, riêng {failed} file không xóa được.",
        )

    def _on_close(self):
        if self.running:
            if not messagebox.askyesno(APP_NAME, "Đang xử lý dở. Thoát ứng dụng bây giờ?"):
                return
            CANCEL_EVENT.set()
        self._save_settings()
        was_running = self.running
        self.destroy()
        if was_running:
            os._exit(0)  # dừng luôn các tiến trình OCR con


if __name__ == "__main__":
    # Bắt buộc khi đóng gói PyInstaller trên Windows, nếu không mỗi tiến trình
    # con sẽ mở thêm một cửa sổ ứng dụng.
    mp.freeze_support()
    if sys.platform == "win32":
        try:  # chữ nét hơn trên màn hình độ phân giải cao
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()