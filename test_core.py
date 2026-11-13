import io
import tempfile
from contextlib import redirect_stdout
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

import numpy as np
import pandas as pd
import fitz
import pdfplumber
from fitz import Document

from app import (
    DocumentInfo,
    confirm_output_path,
    extract_birth_date_candidate,
    extract_birth_date_from_pages,
    extract_birth_dates,
    alternate_norm_date,
    ids_match,
    is_plausible_driving_dob,
    is_valid_vietnamese_id,
    inspect_pdf,
    match_document,
    norm_date,
    norm_text,
    ocr_page,
    process,
    select_id_candidate,
)


def test_normalization():
    assert norm_text("Bùi Quang Tiệp") == "bui quang tiep"
    assert norm_date("8.4.2025") == "08/04/2025"
    assert norm_date("19760804") == "04/08/1976"
    assert norm_date(19760804) == "04/08/1976"
    assert norm_date("19761340") == ""
    assert alternate_norm_date(datetime(2000, 6, 10)) == "06/10/2000"
    assert ids_match("145892756", "145892756")
    assert ids_match("12345678", "012345678")


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
                assert image.shape[0] == round(page.rect.height * 3)
                assert image.shape[1] == round(page.rect.width * 3)
                return [
                    ([], ("CCCD 034202005650", 0.8)),
                    ([], ("23/10/2002", 0.6)),
                    ([], (" ", 1.0)),
                ]

        with pdfplumber.open(pdf_path) as pdf, patch("app.get_deepdoc_ocr", return_value=FakeOCR()):
            rendered = np.array(pdf.pages[0].to_image(resolution=216).annotated)
            assert tuple(rendered.shape) == (
                round(page.rect.height * 3),
                round(page.rect.width * 3),
                3,
            )
            text, confidence, lines = ocr_page(pdf.pages[0])

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
    lines = [("24/10/1999", 92.0), *[("Địa chỉ:", 91.0)] * 12,
             ("Ngày, tháng, năm sinh:", 96.0)]

    assert extract_birth_dates(lines) == set()


def test_nearby_birth_date_beats_higher_confidence_distant_date():
    lines = [
        ("Date of birth:", 87.5),
        ("15/08/1966", 75.0),
        *[("unrelated field", 85.0)] * 9,
        ("20/11/1990", 99.0),
    ]

    assert extract_birth_dates(lines, today=date(2026, 11, 13)) == {"15/08/1966"}


def test_wide_birth_date_fallback_is_flagged_for_manual_review():
    candidate = extract_birth_date_candidate(
        [
            ("Date of birth:", 87.5),
            *[("unrelated field", 85.0)] * 10,
            ("15/08/1966", 90.0),
        ],
        today=date(2026, 11, 13),
    )

    assert candidate == ("15/08/1966", 90.0, False, False, True)


def test_nearby_birth_date_on_another_page_beats_wide_fallback():
    pages = [
        [
            ("Date of birth:", 87.5),
            *[("unrelated field", 85.0)] * 10,
            ("20/11/1990", 99.0),
        ],
        [("Ngày sinh", 89.3), ("15/08/1966", 75.0)],
    ]

    assert extract_birth_date_from_pages(
        pages, today=date(2026, 11, 13)
    ) == ("15/08/1966", 75.0, False, False, False)


def test_birth_date_search_does_not_cross_pdf_pages():
    pages = [
        [("Date of birth:", 90.0)],
        [("unrelated", 90.0), ("15/08/1966", 95.0)],
    ]

    assert extract_birth_date_from_pages(
        pages, today=date(2026, 11, 13)
    ) == (None, None, False, False, False)


def test_vietnamese_id_format_filter_and_score_selection():
    assert is_valid_vietnamese_id("034202005650")
    assert is_valid_vietnamese_id("001202005650")
    assert is_valid_vietnamese_id("096202005650")
    assert not is_valid_vietnamese_id("000202005650")
    assert not is_valid_vietnamese_id("097202005650")
    assert is_valid_vietnamese_id("123456789")
    assert not is_valid_vietnamese_id("034250005650")
    assert not is_valid_vietnamese_id("1234567890")

    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        lines = [
            ("CCCD 034202005650", 68.0),
            ("CCCD 034202005651", 92.0),
            ("CCCD 999202005650", 99.0),
        ]
        with patch(
            "app.ocr_page",
            return_value=(" ".join(text for text, _ in lines), 85.0, lines),
        ):
            info = inspect_pdf(pdf_path)

        assert info.ids == {"034202005651"}
        assert info.id_confidence == 92.0


def test_id_selection_prefers_highest_confidence_candidate():
    assert select_id_candidate(
        [("001204017170", 83.2), ("001204047970", 74.7)],
    ) == ("001204017170", 83.2)


def test_id_selection_prioritizes_12_digits_over_9_digits():
    assert select_id_candidate(
        [("123456789", 99.0), ("001204047970", 61.0)],
    ) == ("001204047970", 61.0)


def test_id_selection_rejects_low_confidence_ocr_noise():
    assert select_id_candidate([("077077000000", 52.6)]) == (None, None)


def test_split_vietnamese_id_uses_mean_line_confidence():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        lines = [("CCCD: 034202", 80.0), ("005650", 60.0)]
        with patch(
            "app.ocr_page",
            return_value=("CCCD: 034202 005650", 70.0, lines),
        ):
            info = inspect_pdf(pdf_path)

        assert info.ids == {"034202005650"}
        assert info.id_confidence == 70.0


def test_driving_dob_plausibility_bounds():
    today = date(2026, 6, 18)
    assert is_plausible_driving_dob("18/06/2008", today)
    assert not is_plausible_driving_dob("19/06/2008", today)
    assert is_plausible_driving_dob("18/06/1926", today)
    assert not is_plausible_driving_dob("17/06/1925", today)
    assert not is_plausible_driving_dob("19/06/2026", today)


def test_birth_date_selects_highest_confidence_plausible_candidate():
    lines = [
        ("Ngày, tháng, năm sinh:", 96.0),
        ("23/10/2002", 75.0),
        ("Ngày sinh:", 90.0),
        ("11/02/1990", 92.0),
        ("Ngày sinh:", 99.0),
        ("19/06/2008", 99.0),
    ]

    assert extract_birth_dates(lines, today=date(2026, 6, 18)) == {"11/02/1990"}


def test_split_birth_date_uses_mean_line_confidence():
    value, confidence, out_of_range, ambiguous, wide = extract_birth_date_candidate(
        [
            ("Ngày, tháng, năm sinh:", 96.0),
            ("24/10/", 80.0),
            ("1999", 90.0),
        ],
        today=date(2026, 6, 18),
    )

    assert value == "24/10/1999"
    assert confidence == 85.0
    assert not out_of_range
    assert not ambiguous
    assert not wide


def test_inspect_pdf_keeps_selected_birth_date_and_score():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        lines = [
            ("Ngày, tháng, năm sinh:", 96.0),
            ("23/10/2002", 75.0),
            ("Ngày sinh:", 90.0),
            ("11/02/1990", 92.0),
        ]
        with patch(
            "app.ocr_page",
            return_value=(" ".join(text for text, _ in lines), 85.0, lines),
        ):
            info = inspect_pdf(pdf_path)

        assert info.birth_dates == {"11/02/1990"}
        assert info.birth_date_confidence == 92.0
        assert not info.birth_date_out_of_range
        assert not info.birth_date_ambiguous


def test_inspect_pdf_marks_wide_dob_fallback_for_manual_review():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        lines = [
            ("Date of birth:", 87.5),
            *[("unrelated field", 85.0)] * 10,
            ("15/08/1966", 90.0),
        ]
        with patch(
            "app.ocr_page",
            return_value=(" ".join(text for text, _ in lines), 85.0, lines),
        ):
            info = inspect_pdf(pdf_path)

        assert info.birth_dates == {"15/08/1966"}
        assert info.birth_date_wide_search


def test_birth_date_reports_nearby_out_of_range_candidate():
    value, confidence, out_of_range, ambiguous, wide = extract_birth_date_candidate(
        [
            ("Ngày, tháng, năm sinh:", 96.0),
            ("19/06/2008", 99.0),
        ],
        today=date(2026, 6, 18),
    )

    assert value is None
    assert confidence is None
    assert out_of_range
    assert not ambiguous
    assert not wide


def test_tied_plausible_birth_dates_remain_ambiguous():
    value, confidence, out_of_range, ambiguous, wide = extract_birth_date_candidate(
        [
            ("Ngày, tháng, năm sinh:", 96.0),
            ("23/10/2002 24/10/2002", 90.0),
        ],
        today=date(2026, 6, 18),
    )

    assert value is None
    assert confidence is None
    assert not out_of_range
    assert ambiguous
    assert not wide


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


def test_inspect_pdf_uses_ocr_id_score_from_pdfplumber_render():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)

        class FakeOCR:
            def __call__(self, image):
                return [([], ("CCCD 033181011361", 0.9))]

        with patch("app.get_deepdoc_ocr", return_value=FakeOCR()):
            info = inspect_pdf(pdf_path)

        assert info.ids == {"033181011361"}
        assert info.id_confidence == 90.0
        assert info.used_ocr


def test_inspect_pdf_prefers_12_digit_content_id_over_filename_id():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "Bui Quang Tiep 145892756.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        with patch(
            "app.ocr_page",
            return_value=(
                "CCCD 034202005650",
                90.0,
                [("CCCD 034202005650", 90.0)],
            ),
        ):
            info = inspect_pdf(pdf_path)

        assert info.ids == {"034202005650"}
        assert info.id_confidence == 90.0
        assert info.id_from_pdf


def test_inspect_pdf_uses_filename_9_digit_id_when_pdf_has_no_12_digit_id():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "Bui Quang Tiep 145892756.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        with patch(
            "app.ocr_page",
            return_value=("unreadable", 80.0, [("unreadable", 80.0)]),
        ):
            info = inspect_pdf(pdf_path)

        assert info.ids == {"145892756"}
        assert info.id_confidence is None
        assert not info.id_from_pdf


def test_match_document_does_not_use_filename_name_without_pdf_id():
    pdf_path = Path("Nguyen Van Thuan.pdf")
    doc = DocumentInfo(
        pdf_path, "", set(), set(), set(), "nguyen van thuan", 82.1, True
    )

    assert match_document("Nguyễn Văn Thuận", "033093009330", [doc]) == (None, 0.0)
    filename_id_doc = DocumentInfo(
        pdf_path, "", {"033093009330"}, set(), set(), "nguyen van thuan",
        82.1, True, id_from_pdf=True, filename_ids={"145892756"},
    )
    assert match_document("Nguyễn Văn Thuận", "145892756", [filename_id_doc]) == (
        filename_id_doc,
        1.0,
    )


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


def test_pdf_with_exact_id_is_not_assigned_to_multiple_rows():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        source = root / "danh_sach.xlsx"
        out = root / "ket_qua.xlsx"
        pdf_dir = root / "pdf"
        pdf_dir.mkdir()
        document = Document()
        document.new_page().insert_text((72, 72), "CCCD 034202005650")
        document.save(pdf_dir / "Bui Quang Tiep 034202005650.pdf")
        pd.DataFrame([
            {"Họ và tên": "Bùi Quang Tiệp", "Số CCCD": "034202005650"},
            {"Họ và tên": "Bùi Quang Tiệp", "Số CCCD": "034202005650"},
        ]).to_excel(source, index=False)

        counts, _ = process(source, pdf_dir, out, lambda *args: None)
        result = pd.read_excel(out)

        assert counts == {"OK": 1, "Không tìm thấy hồ sơ": 1}
        assert result.loc[0, "Đường dẫn tài liệu"].endswith(".pdf")
        assert pd.isna(result.loc[1, "Đường dẫn tài liệu"])


def test_filename_id_fallback_is_matched_but_marked_for_review():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        source = root / "danh_sach.xlsx"
        out = root / "ket_qua.xlsx"
        pdf_dir = root / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "Bui Quang Tiep 145892756.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        pd.DataFrame([{
            "Họ và tên": "Bùi Quang Tiệp",
            "Số CCCD": "145892756",
        }]).to_excel(source, index=False)
        info = DocumentInfo(
            pdf_path, "CCCD 033181011361", {"033181011361"}, set(), set(),
            "bui quang tiep", 90.0, True,
            id_from_pdf=True, filename_ids={"145892756"},
        )

        with patch("app.inspect_pdf", return_value=info):
            counts, _ = process(source, pdf_dir, out, lambda *args: None)

        result = pd.read_excel(out)
        assert counts == {"Cần kiểm tra": 1}
        assert result.loc[0, "Đường dẫn tài liệu"].endswith(".pdf")
        assert "Ghép theo ID trong tên file" in result.loc[0, "Chi tiết sai lệch"]
        assert "ID đọc từ PDF" in result.loc[0, "Chi tiết sai lệch"]


def test_fuzzy_filename_match_selects_highest_similarity_row_and_marks_wrong():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        source = root / "danh_sach.xlsx"
        out = root / "ket_qua.xlsx"
        pdf_dir = root / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "Nguyen Van A 123456789.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        pd.DataFrame([
            {"Họ và tên": "Nguyễn Văn An", "Số CCCD": "123456780"},
            {"Họ và tên": "Nguyễn Văn A", "Số CCCD": "123456781"},
        ]).to_excel(source, index=False)
        info = DocumentInfo(
            pdf_path, "", {"034202005650"}, set(), set(),
            "nguyen van a", 20.0, True,
            id_from_pdf=True, filename_ids={"123456789"},
        )

        with patch("app.inspect_pdf", return_value=info):
            counts, _ = process(source, pdf_dir, out, lambda *args: None)

        result = pd.read_excel(out)
        assert counts == {"Sai thông tin": 1, "Không tìm thấy hồ sơ": 1}
        assert pd.isna(result.loc[0, "Đường dẫn tài liệu"])
        assert result.loc[1, "Đường dẫn tài liệu"].endswith(".pdf")
        assert result.loc[1, "Trạng thái đối soát"] == "Sai thông tin"
        assert "độ khớp tên 100%" in result.loc[1, "Chi tiết sai lệch"]
        assert "ID 89%" in result.loc[1, "Chi tiết sai lệch"]


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
        assert "đảo ngày/tháng" in result.loc[0, "Chi tiết sai lệch"]


def test_day_month_swapped_birth_date_is_review_not_mismatch():
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
            "Ngày sinh": datetime(2000, 6, 10),
        }]).to_excel(source, index=False)
        info = DocumentInfo(
            pdf_path, "", {"034202005650"}, set(), {"06/10/2000"},
            "bui quang tiep", 92.0, True, id_from_pdf=True,
        )

        with patch("app.inspect_pdf", return_value=info):
            counts, _ = process(source, pdf_dir, out, lambda *args: None)

        result = pd.read_excel(out)
        assert counts == {"Cần kiểm tra": 1}
        assert result.loc[0, "Trạng thái đối soát"] == "Cần kiểm tra"
        assert "đảo ngày/tháng" in result.loc[0, "Chi tiết sai lệch"]
        assert "Sai thông tin" not in result.loc[0, "Trạng thái đối soát"]


def test_mismatch_status_requires_pdf_confidence_of_at_least_80():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        source = root / "danh_sach.xlsx"
        pdf_dir = root / "pdf"
        pdf_dir.mkdir()
        pdf_path = pdf_dir / "Bui Quang Tiep 034202005650.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        pd.DataFrame([{
            "Họ và tên": "Bùi Quang Tiệp",
            "Số CCCD": "034202005650",
            "Ngày sinh": "02/04/1999",
        }]).to_excel(source, index=False)

        for confidence, expected in ((79.9, "Cần kiểm tra"), (80.0, "Sai thông tin")):
            out = root / f"ket_qua_{confidence}.xlsx"
            info = DocumentInfo(
                pdf_path, "", {"034202005650"}, set(), {"05/05/1980"},
                "bui quang tiep", confidence, True, id_from_pdf=True,
            )
            with patch("app.inspect_pdf", return_value=info):
                counts, _ = process(source, pdf_dir, out, lambda *args: None)

            result = pd.read_excel(out)
            assert counts == {expected: 1}
            assert result.loc[0, "Trạng thái đối soát"] == expected


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
        info = DocumentInfo(
            pdf_path, "unreadable", {"034202005650"}, set(), set(),
            "bui quang tiep", 20, True, id_from_pdf=True,
        )

        with patch("app.inspect_pdf", return_value=info):
            process(source, pdf_dir, out, lambda *args: None)

        result = pd.read_excel(out)
        detail = result.loc[0, "Chi tiết sai lệch"]
        assert "Điểm OCR thấp" in detail
        assert "CCCD" not in detail
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
            pdf_path, "recognized text", {"034202005650"}, set(), set(),
            "bui quang tiep", None, True, id_from_pdf=True,
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
            id_from_pdf=True,
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
    test_vietnamese_id_format_filter_and_score_selection()
    test_id_selection_prefers_highest_confidence_candidate()
    test_id_selection_prioritizes_12_digits_over_9_digits()
    test_split_vietnamese_id_uses_mean_line_confidence()
    test_driving_dob_plausibility_bounds()
    test_birth_date_selects_highest_confidence_plausible_candidate()
    test_split_birth_date_uses_mean_line_confidence()
    test_inspect_pdf_keeps_selected_birth_date_and_score()
    test_birth_date_reports_nearby_out_of_range_candidate()
    test_tied_plausible_birth_dates_remain_ambiguous()
    test_inspect_pdf_aggregates_ocr_confidence()
    test_inspect_pdf_uses_ocr_id_score_from_pdfplumber_render()
    test_inspect_pdf_prefers_12_digit_content_id_over_filename_id()
    test_inspect_pdf_uses_filename_9_digit_id_when_pdf_has_no_12_digit_id()
    test_match_document_does_not_use_filename_name_without_pdf_id()
    test_existing_output_can_be_overwritten_or_renamed()
    test_end_to_end_text_pdf()
    test_pdf_with_exact_id_is_not_assigned_to_multiple_rows()
    test_filename_id_fallback_is_matched_but_marked_for_review()
    test_fuzzy_filename_match_selects_highest_similarity_row_and_marks_wrong()
    test_swapped_dob_format_is_warning_not_mismatch()
    test_day_month_swapped_birth_date_is_review_not_mismatch()
    test_mismatch_status_requires_pdf_confidence_of_at_least_80()
    test_low_ocr_confidence_names_uncertain_fields()
    test_missing_deepdoc_confidence_requires_manual_review()
    test_ocr_confidence_is_written_to_results()
    print("All tests passed")
