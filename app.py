from __future__ import annotations

import csv
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
from queue import Empty, Queue
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


APP_NAME = "Đối soát hồ sơ"
OUTPUT_COLUMNS = [
    "Trạng thái đối soát", "Chi tiết sai lệch", "Độ tin cậy OCR",
    "Đường dẫn tài liệu", "Ghi chú kiểm tra thủ công",
]


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
    )


def norm_date(value) -> str:
    if value is None or pd.isna(value) or value == "":
        return ""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.strftime("%d/%m/%Y")
    s = str(value).strip()
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
    if value is None or pd.isna(value) or isinstance(value, (datetime, date, pd.Timestamp)):
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


def find_tesseract() -> str | None:
    candidates = [
        resource_dir() / "tesseract" / "tesseract.exe",
        Path(sys.executable).resolve().parent / "tesseract" / "tesseract.exe",
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    ]
    return next((str(p) for p in candidates if p.exists()), shutil.which("tesseract"))


def open_folder(path: Path):
    """Mở thư mục kết quả trên Windows, macOS hoặc Linux."""
    folder = str(path.resolve())
    if sys.platform == "win32":
        os.startfile(folder)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])


@lru_cache(maxsize=8)
def tesseract_languages(tesseract: str) -> set[str]:
    env = os.environ.copy()
    bundled_tessdata = Path(tesseract).parent / "tessdata"
    if bundled_tessdata.exists():
        env["TESSDATA_PREFIX"] = str(bundled_tessdata)
    run = subprocess.run(
        [tesseract, "--list-langs"],
        capture_output=True,
        text=True,
        env=env,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if run.returncode:
        raise RuntimeError(run.stderr.strip() or "Không thể kiểm tra ngôn ngữ Tesseract")
    return {
        line.strip()
        for line in run.stdout.splitlines()
        if re.fullmatch(r"[a-z0-9_]+", line.strip())
    }


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


def read_ocr_tsv(path: Path) -> tuple[str, float]:
    tsv = pd.read_csv(
        path,
        sep="\t",
        keep_default_na=False,
        quoting=csv.QUOTE_NONE,
    )
    words = tsv.loc[tsv["text"].astype(str).str.strip() != "", "text"].astype(str).tolist()
    conf = pd.to_numeric(tsv.get("conf", pd.Series(dtype=float)), errors="coerce")
    conf = conf[conf >= 0]
    return " ".join(words), float(conf.mean()) if len(conf) else 0.0


def ocr_page(page: fitz.Page, tesseract: str) -> tuple[str, float]:
    pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), colorspace=fitz.csGRAY, alpha=False)
    with tempfile.TemporaryDirectory(prefix="hoso_ocr_") as td:
        image = Path(td) / "page.png"
        out = Path(td) / "result"
        pix.save(str(image))
        lang = "vie+eng" if "vie" in tesseract_languages(tesseract) else "eng"
        print(f"Đang OCR trang PDF bằng Tesseract ({lang})...")
        cmd = [tesseract, str(image), str(out), "-l", lang, "--psm", "6", "tsv"]
        env = os.environ.copy()
        bundled_tessdata = Path(tesseract).parent / "tessdata"
        if bundled_tessdata.exists():
            env["TESSDATA_PREFIX"] = str(bundled_tessdata)
        run = subprocess.run(
            cmd, capture_output=True, text=True, env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if run.returncode:
            raise RuntimeError(run.stderr.strip() or "OCR thất bại")
        return read_ocr_tsv(Path(str(out) + ".tsv"))


@dataclass
class DocumentInfo:
    path: Path
    text: str
    ids: set[str]
    dates: set[str]
    birth_dates: set[str]
    name_hint: str
    confidence: float
    used_ocr: bool
    error: str = ""


def inspect_pdf(path: Path, tesseract: str | None) -> DocumentInfo:
    chunks, scores, used_ocr, error = [], [], False, ""
    try:
        with fitz.open(path) as doc:
            for page in list(doc)[:8]:
                text = page.get_text("text").strip()
                if len(re.sub(r"\s", "", text)) >= 40:
                    chunks.append(text)
                    scores.append(100.0)
                elif tesseract:
                    used_ocr = True
                    ocr_text, conf = ocr_page(page, tesseract)
                    chunks.append(ocr_text)
                    scores.append(conf)
                else:
                    error = "PDF scan nhưng chưa có bộ OCR Tesseract"
    except Exception as exc:
        error = str(exc)
    raw = "\n".join(chunks)
    combined = path.stem + "\n" + raw
    ids = set()
    for candidate in re.findall(r"(?<!\d)(?:\d[ .-]?){9,12}(?!\d)", combined):
        digits = norm_digits(candidate)
        if len(digits) in (9, 12):
            ids.add(digits)
    dates = set()
    for d, m, y in re.findall(r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?!\d)", raw):
        value = norm_date(f"{d}/{m}/{y}")
        if value:
            dates.add(value)
    birth_dates = set()
    flat = re.sub(r"\s+", " ", raw)
    dob_pattern = r"(?:ng[aà]y(?: th[aá]ng n[aă]m)? sinh|date of birth)[^\d]{0,35}(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})"
    for found in re.findall(dob_pattern, flat, flags=re.IGNORECASE):
        value = norm_date(found)
        if value:
            birth_dates.add(value)
    name_hint = re.sub(r"\b\d[\d ._-]*\b", " ", path.stem)
    name_hint = norm_text(name_hint)
    confidence = sum(scores) / len(scores) if scores else 0.0
    return DocumentInfo(path, raw, ids, dates, birth_dates, name_hint, confidence, used_ocr, error)


ALIASES = {
    "name": ["ho va ten", "ho ten", "ten khach hang", "ten nguoi", "customer name"],
    "id": ["so cccd", "cccd", "cmnd", "so cmnd", "so chung minh", "chung minh nhan dan", "can cuoc", "identity"],
    "dob": ["ngay sinh", "nam sinh", "date of birth", "dob"],
}


def detect_columns(columns) -> dict[str, str | None]:
    normalized = {c: norm_text(c) for c in columns}
    result = {}
    for field, aliases in ALIASES.items():
        result[field] = next((c for c, n in normalized.items() if n in aliases), None)
        if not result[field]:
            result[field] = next((c for c, n in normalized.items() if any(a in n for a in aliases)), None)
    return result


def match_document(name: str, citizen_id: str, docs: list[DocumentInfo]) -> tuple[DocumentInfo | None, float]:
    if citizen_id:
        exact = [
            d for d in docs
            if any(ids_match(citizen_id, candidate) for candidate in d.ids)
            or ids_match(citizen_id, norm_digits(d.path.stem))
        ]
        if exact:
            return exact[0], 1.0
    n = norm_text(name)
    if not n:
        return None, 0.0
    # Nếu cả Excel và PDF đều có CCCD nhưng khác nhau thì tuyệt đối không ghép
    # chỉ vì tên gần giống: đây là lỗi nguy hiểm nhất khi danh sách dài.
    candidates = [
        d for d in docs
        if not citizen_id or not d.ids or any(ids_match(citizen_id, candidate) for candidate in d.ids)
    ]
    ranked = sorted(((SequenceMatcher(None, n, d.name_hint).ratio(), d) for d in candidates), reverse=True, key=lambda x: x[0])
    return (ranked[0][1], ranked[0][0]) if ranked and ranked[0][0] >= 0.67 else (None, ranked[0][0] if ranked else 0.0)


def process(excel_path: Path, pdf_dir: Path, output_path: Path, progress):
    df = pd.read_excel(excel_path, dtype=object)
    cols = detect_columns(df.columns)
    if not cols["name"] or not cols["id"]:
        raise ValueError("Không tự nhận ra cột Họ tên/CCCD. Hãy đổi tiêu đề cột thành 'Họ và tên' và 'Số CCCD'.")
    pdfs = sorted(pdf_dir.rglob("*.pdf"))
    if not pdfs:
        raise ValueError("Thư mục đã chọn không có file PDF.")
    tess = find_tesseract()
    docs = []
    for i, path in enumerate(pdfs, 1):
        progress("scan", i, len(pdfs), f"Đang đọc {path.name}")
        docs.append(inspect_pdf(path, tess))

    results, used = [], set()
    for i, (_, row) in enumerate(df.iterrows(), 1):
        progress("match", i, len(df), f"Đang đối chiếu dòng {i}")
        name = str(row.get(cols["name"], "") or "")
        cid = norm_digits(row.get(cols["id"], ""))
        raw_dob = row.get(cols["dob"], "") if cols["dob"] else ""
        dob = norm_date(raw_dob)
        alternate_dob = alternate_norm_date(raw_dob)
        doc, score = match_document(name, cid, docs)
        if not doc:
            results.append(["Không tìm thấy hồ sơ", "Không tìm thấy PDF khớp CCCD hoặc họ tên", "", "", ""])
            continue
        used.add(doc.path)
        mismatches = []
        format_warnings = []
        if cid and doc.ids and not any(ids_match(cid, candidate) for candidate in doc.ids):
            mismatches.append(f"CCCD Excel {cid} không khớp PDF ({', '.join(sorted(doc.ids))})")
        if raw_dob is not None and not pd.isna(raw_dob) and str(raw_dob).strip() and not dob:
            if alternate_dob and alternate_dob in doc.birth_dates:
                format_warnings.append(
                    f"Ngày sinh Excel có thể đang theo định dạng tháng/ngày; "
                    f"đọc theo tháng/ngày là {alternate_dob}, khớp với PDF"
                )
            else:
                format_warnings.append("Không nhận diện được định dạng ngày sinh trong Excel; cần kiểm tra lại ngày/tháng")
        elif dob and doc.birth_dates and dob not in doc.birth_dates:
            if alternate_dob and alternate_dob in doc.birth_dates:
                format_warnings.append(
                    f"Ngày sinh Excel có thể đang theo định dạng tháng/ngày; "
                    f"đọc theo ngày/tháng là {dob}, còn PDF khớp {alternate_dob}"
                )
            else:
                mismatches.append(f"Ngày sinh Excel {dob} không khớp PDF ({', '.join(sorted(doc.birth_dates))})")
        ocr_warning = ""
        if doc.used_ocr and doc.confidence < 50:
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
            status, detail = "Sai thông tin", "; ".join([*mismatches, *format_warnings, ocr_warning])
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
        results.append([status, detail, round(doc.confidence, 1) if doc.used_ocr else 100, str(doc.path.resolve()), ""])

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
    return counts, len(pdfs), bool(tess)


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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("760x500")
        self.minsize(680, 440)
        self.excel = tk.StringVar()
        self.folder = tk.StringVar()
        self.output = tk.StringVar()
        self.queue = Queue()
        self._build()

    def _build(self):
        root = ttk.Frame(self, padding=24)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="ĐỐI SOÁT HỒ SƠ", font=("Segoe UI", 19, "bold")).pack(anchor="w")
        ttk.Label(root, text="Chọn danh sách Excel và thư mục PDF, sau đó bấm Bắt đầu kiểm tra.").pack(anchor="w", pady=(4, 22))
        self._picker(root, "File Excel", self.excel, self.pick_excel)
        self._picker(root, "Thư mục PDF", self.folder, self.pick_folder)
        self._picker(root, "File kết quả", self.output, self.pick_output)
        self.start = ttk.Button(root, text="Bắt đầu kiểm tra", command=self.run)
        self.start.pack(anchor="w", pady=(18, 12), ipadx=12, ipady=6)
        self.bar = ttk.Progressbar(root, mode="determinate")
        self.bar.pack(fill="x")
        self.status = ttk.Label(root, text="Sẵn sàng", wraplength=700)
        self.status.pack(anchor="w", pady=8)
        self.summary = ttk.Label(root, text="", font=("Segoe UI", 11, "bold"))
        self.summary.pack(anchor="w", pady=8)

    def _picker(self, parent, label, variable, command):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=5)
        ttk.Label(frame, text=label, width=14).pack(side="left")
        ttk.Entry(frame, textvariable=variable).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(frame, text="Chọn...", command=command).pack(side="right")

    def pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.excel.set(p)
            self.output.set(str(Path(p).with_name(Path(p).stem + "_ket_qua.xlsx")))

    def pick_folder(self):
        if p := filedialog.askdirectory():
            self.folder.set(p)

    def pick_output(self):
        if p := filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]):
            self.output.set(p)

    def run(self):
        if not self.excel.get() or not self.folder.get() or not self.output.get():
            messagebox.showwarning(APP_NAME, "Vui lòng chọn đủ file Excel, thư mục PDF và file kết quả.")
            return
        output_path = Path(self.output.get())
        if output_path.resolve() == Path(self.excel.get()).resolve():
            messagebox.showerror(APP_NAME, "File kết quả không được trùng với file Excel nguồn.")
            return
        output_path = confirm_output_path(output_path)
        if output_path is None:
            return
        self.output.set(str(output_path))
        self.start.config(state="disabled")
        self.summary.config(text="")
        threading.Thread(target=self._worker, daemon=True).start()
        self.after(100, self._poll)

    def _worker(self):
        try:
            result = process(Path(self.excel.get()), Path(self.folder.get()), Path(self.output.get()),
                             lambda phase, n, total, msg: self.queue.put(("progress", n, total, msg)))
            self.queue.put(("done", result))
        except Exception as exc:
            self.queue.put(("error", str(exc)))

    def _poll(self):
        try:
            while True:
                event = self.queue.get_nowait()
                if event[0] == "progress":
                    _, n, total, msg = event
                    self.bar["maximum"], self.bar["value"] = total, n
                    self.status.config(text=msg)
                elif event[0] == "done":
                    counts, pdf_count, has_ocr = event[1]
                    text = " | ".join(f"{k}: {v}" for k, v in counts.items())
                    self.summary.config(text=text)
                    self.status.config(text=f"Hoàn tất {pdf_count} PDF" + ("" if has_ocr else " (máy chưa có OCR; PDF scan sẽ cần kiểm tra)"))
                    self.start.config(state="normal")
                    if messagebox.askyesno(APP_NAME, "Đã xuất file kết quả. Mở thư mục chứa file?"):
                        open_folder(Path(self.output.get()).parent)
                    return
                elif event[0] == "error":
                    self.start.config(state="normal")
                    messagebox.showerror(APP_NAME, event[1])
                    return
        except Empty:
            self.after(100, self._poll)


if __name__ == "__main__":
    App().mainloop()
