import io
import subprocess
import tempfile
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from fitz import Document

from app import (
    DocumentInfo,
    confirm_output_path,
    norm_date,
    norm_text,
    process,
    read_ocr_tsv,
    tesseract_languages,
)


def test_normalization():
    assert norm_text("Bùi Quang Tiệp") == "bui quang tiep"
    assert norm_date("8.4.2025") == "08/04/2025"


def test_tesseract_languages_reads_installed_language_list():
    tesseract_languages.cache_clear()
    result = subprocess.CompletedProcess(
        args=["tesseract", "--list-langs"],
        returncode=0,
        stdout='List of available languages in "/share/tessdata/":\neng\nvie\n',
        stderr="",
    )
    with patch("app.subprocess.run", return_value=result) as run:
        assert tesseract_languages("/usr/bin/tesseract") == {"eng", "vie"}
        run.assert_called_once()
    tesseract_languages.cache_clear()


def test_ocr_tsv_with_unclosed_quote_is_read_without_csv_error():
    with tempfile.TemporaryDirectory() as td:
        tsv_path = Path(td) / "ocr.tsv"
        lines = [
            "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext"
        ]
        for row in range(1, 38):
            text = '"unclosed' if row == 37 else "word"
            lines.append(f"5\t1\t1\t1\t{row}\t1\t0\t0\t10\t10\t90\t{text}")
        tsv_path.write_text("\n".join(lines), encoding="utf-8")

        text, confidence = read_ocr_tsv(tsv_path)

        assert text.endswith('"unclosed')
        assert confidence == 90


def test_existing_output_can_be_overwritten_or_renamed():
    with tempfile.TemporaryDirectory() as td:
        existing = Path(td) / "result.xlsx"
        existing.touch()
        renamed = Path(td) / "result_moi.xlsx"

        with patch("app.messagebox.askyesnocancel", return_value=True):
            assert confirm_output_path(existing) == existing

        with (
            patch("app.messagebox.askyesnocancel", return_value=False),
            patch("app.filedialog.asksaveasfilename", return_value=str(renamed)),
        ):
            assert confirm_output_path(existing) == renamed

        with patch("app.messagebox.askyesnocancel", return_value=None):
            assert confirm_output_path(existing) is None


def test_end_to_end_text_pdf():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        source = root / "danh_sach.xlsx"
        out = root / "ket_qua.xlsx"
        pdf_dir = root / "pdf"
        pdf_dir.mkdir()
        doc = Document()
        page = doc.new_page()
        page.insert_text((72, 72), "Bui Quang Tiep  CCCD: 034202005650  Ngay sinh: 23/10/2002")
        doc.save(pdf_dir / "Bui Quang Tiep - 034202005650.pdf")
        unmatched = Document()
        unmatched.new_page()
        unmatched.save(pdf_dir / "unmatched.pdf")
        pd.DataFrame([{
            "Họ và tên": "Bùi Quang Tiệp",
            "Số CCCD": 34202005650,
            "Ngày sinh": "23/10/2002",
        }]).to_excel(source, index=False)
        printed = io.StringIO()
        with redirect_stdout(printed):
            counts, total, _ = process(source, pdf_dir, out, lambda *args: None)
        assert total == 2
        assert counts == {"OK": 1}
        result = pd.read_excel(out)
        assert result.loc[0, "Trạng thái đối soát"] == "OK"
        assert result.loc[0, "Đường dẫn tài liệu"].endswith(".pdf")
        assert "unmatched.pdf" in printed.getvalue()


def test_swapped_dob_format_is_warning_not_mismatch():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        source = root / "danh_sach.xlsx"
        out = root / "ket_qua.xlsx"
        pdf_dir = root / "pdf"
        pdf_dir.mkdir()
        pd.DataFrame([{
            "Họ và tên": "Bùi Quang Tiệp",
            "Số CCCD": "034202005650",
            "Ngày sinh": "10/23/2002",
        }]).to_excel(source, index=False)
        doc = Document()
        page = doc.new_page()
        page.insert_text((72, 72), "Bui Quang Tiep  CCCD: 034202005650  Ngay sinh: 23/10/2002")
        doc.save(pdf_dir / "Bui Quang Tiep - 034202005650.pdf")

        counts, _, _ = process(source, pdf_dir, out, lambda *args: None)
        result = pd.read_excel(out)
        assert counts == {"Cần kiểm tra": 1}
        assert result.loc[0, "Trạng thái đối soát"] == "Cần kiểm tra"
        assert "định dạng tháng/ngày" in result.loc[0, "Chi tiết sai lệch"]


def test_low_ocr_confidence_names_uncertain_fields():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        source = root / "danh_sach.xlsx"
        out = root / "ket_qua.xlsx"
        pdf_dir = root / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "Bui Quang Tiep.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        pd.DataFrame([{
            "Họ và tên": "Bùi Quang Tiệp",
            "Số CCCD": "034202005650",
            "Ngày sinh": "23/10/2002",
        }]).to_excel(source, index=False)
        info = DocumentInfo(pdf_path, "unreadable", set(), set(), set(), "bui quang tiep", 20, True)

        with patch("app.inspect_pdf", return_value=info):
            process(source, pdf_dir, out, lambda *args: None)

        result = pd.read_excel(out)
        detail = result.loc[0, "Chi tiết sai lệch"]
        assert "Điểm OCR thấp" in detail
        assert "CCCD" in detail
        assert "ngày sinh" in detail
        assert "họ tên" in detail


if __name__ == "__main__":
    test_normalization()
    test_tesseract_languages_reads_installed_language_list()
    test_ocr_tsv_with_unclosed_quote_is_read_without_csv_error()
    test_existing_output_can_be_overwritten_or_renamed()
    test_end_to_end_text_pdf()
    test_swapped_dob_format_is_warning_not_mismatch()
    test_low_ocr_confidence_names_uncertain_fields()
    print("All tests passed")
