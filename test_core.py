import io
import tempfile
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import fitz
from fitz import Document

from app import (
    DocumentInfo,
    confirm_output_path,
    extract_birth_dates,
    inspect_pdf,
    norm_date,
    norm_text,
    ocr_page,
    process,
)


def test_normalization():
    assert norm_text("Bùi Quang Tiệp") == "bui quang tiep"
    assert norm_date("8.4.2025") == "08/04/2025"
    assert norm_date("19760804") == "04/08/1976"
    assert norm_date(19760804) == "04/08/1976"
    assert norm_date("19761340") == ""


def test_ocr_page_reads_deepdoc_result_text_and_score():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        page = document.new_page()
        document.save(pdf_path)

        class FakeOCR:
            def __call__(self, image):
                assert image.ndim == 3
                assert image.shape[2] == 3
                return [
                    ([], ("CCCD 034202005650", 0.8)),
                    ([], ("23/10/2002", 0.6)),
                    ([], (" ", 1.0)),
                ]

        with fitz.open(pdf_path) as pdf, patch("app.get_deepdoc_ocr", return_value=FakeOCR()):
            text, confidence, lines = ocr_page(pdf[0])

        assert text == "CCCD 034202005650 23/10/2002"
        assert abs(confidence - 72.5925925926) < 0.001
        assert lines == [
            ("CCCD 034202005650", 80.0),
            ("23/10/2002", 60.0),
        ]


def test_extract_birth_date_from_line_before_label():
    lines = [
        ("Họ và tên gọi khác:", 91.0),
        ("05/03/1993", 87.0),
        ("Ngày, tháng, năm sinh:", 94.0),
        ("Dân tộc:", 90.0),
        ("Kinh", 99.0),
    ]

    assert extract_birth_dates(lines) == {"05/03/1993"}


def test_extract_birth_date_when_ocr_label_has_extra_noise():
    lines = [
        ("Ngày, tháng, năm sinh DDte of binh", 82.0),
        ("Nam", 96.0),
        ("24/10/1999", 88.0),
    ]

    assert extract_birth_dates(lines) == {"24/10/1999"}


def test_extract_birth_date_ignores_low_confidence_label_or_date():
    assert extract_birth_dates([
        ("Ngày, tháng, năm sinh:", 42.0),
        ("24/10/1999", 93.0),
    ]) == set()
    assert extract_birth_dates([
        ("Ngày, tháng, năm sinh:", 93.0),
        ("24/10/1999", 49.0),
    ]) == set()


def test_extract_birth_date_ignores_date_outside_nearby_lines():
    lines = [
        ("24/10/1999", 92.0),
        ("Địa chỉ:", 91.0),
        ("Hà Nội", 93.0),
        ("Ngày, tháng, năm sinh:", 96.0),
    ]

    assert extract_birth_dates(lines) == set()


def test_inspect_pdf_aggregates_ocr_confidence():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)

        with patch(
            "app.ocr_page",
            return_value=("recognized text", 86.7, [("recognized text", 86.7)]),
        ):
            info = inspect_pdf(pdf_path)

        assert info.used_ocr
        assert info.confidence == 86.7
        assert info.text == "recognized text"
        assert info.text_lines == [("recognized text", 86.7)]


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
        live_results = []
        scan_events = []

        def capture_result(row, name, citizen_id, result):
            assert not out.exists()
            live_results.append((row, name, citizen_id, result))

        def capture_scan(index, total, path, phase, info):
            assert not out.exists()
            scan_events.append((index, path.name, phase))

        with redirect_stdout(printed):
            counts, total = process(
                source,
                pdf_dir,
                out,
                lambda *args: None,
                capture_result,
                capture_scan,
            )
        assert total == 2
        assert counts == {"OK": 1}
        assert scan_events == [
            (1, "Bui Quang Tiep - 034202005650.pdf", "start"),
            (1, "Bui Quang Tiep - 034202005650.pdf", "done"),
            (2, "unmatched.pdf", "start"),
            (2, "unmatched.pdf", "done"),
        ]
        assert len(live_results) == 1
        assert live_results[0][0:3] == (1, "Bùi Quang Tiệp", "34202005650")
        assert live_results[0][3][0] == "OK"
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

        counts, _ = process(source, pdf_dir, out, lambda *args: None)
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


def test_missing_deepdoc_confidence_requires_manual_review():
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
        }]).to_excel(source, index=False)
        info = DocumentInfo(
            pdf_path, "recognized text", set(), set(), set(),
            "bui quang tiep", None, True,
        )

        with patch("app.inspect_pdf", return_value=info):
            process(source, pdf_dir, out, lambda *args: None)

        result = pd.read_excel(out)
        assert result.loc[0, "Trạng thái đối soát"] == "Cần kiểm tra"
        assert "Không tính được điểm tin cậy OCR" in result.loc[0, "Chi tiết sai lệch"]
        assert pd.isna(result.loc[0, "Độ tin cậy OCR"])


def test_ocr_confidence_is_written_to_results():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        source = root / "danh_sach.xlsx"
        out = root / "ket_qua.xlsx"
        pdf_dir = root / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "Bui Quang Tiep 034202005650.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        pd.DataFrame([{
            "Họ và tên": "Bùi Quang Tiệp",
            "Số CCCD": "034202005650",
        }]).to_excel(source, index=False)
        info = DocumentInfo(
            pdf_path,
            "Bui Quang Tiep CCCD 034202005650",
            {"034202005650"},
            set(),
            set(),
            "bui quang tiep",
            86.7,
            True,
        )

        with patch("app.inspect_pdf", return_value=info):
            process(source, pdf_dir, out, lambda *args: None)

        result = pd.read_excel(out)
        assert result.loc[0, "Độ tin cậy OCR"] == 86.7


if __name__ == "__main__":
    test_normalization()
    test_ocr_page_reads_deepdoc_result_text_and_score()
    test_extract_birth_date_from_line_before_label()
    test_extract_birth_date_when_ocr_label_has_extra_noise()
    test_extract_birth_date_ignores_low_confidence_label_or_date()
    test_extract_birth_date_ignores_date_outside_nearby_lines()
    test_inspect_pdf_aggregates_ocr_confidence()
    test_existing_output_can_be_overwritten_or_renamed()
    test_end_to_end_text_pdf()
    test_swapped_dob_format_is_warning_not_mismatch()
    test_low_ocr_confidence_names_uncertain_fields()
    test_missing_deepdoc_confidence_requires_manual_review()
    test_ocr_confidence_is_written_to_results()
    print("All tests passed")
