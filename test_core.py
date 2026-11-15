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
    extract_ocr_birth_dates,
    alternate_norm_date,
    ids_match,
    is_plausible_driving_dob,
    is_valid_vietnamese_id,
    inspect_pdf,
    match_document,
    norm_date,
    norm_text,
    open_excel_workbook,
    ocr_page,
    process,
    select_id_candidate,
    vietnamese_id_birth_year,
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


def test_excel_reader_engine_is_selected_by_extension():
    with patch("app.pd.ExcelFile") as excel_file:
        open_excel_workbook(Path("list.xls"))
        excel_file.assert_called_once_with(Path("list.xls"), engine="xlrd")

    with patch("app.pd.ExcelFile") as excel_file:
        open_excel_workbook(Path("list.XLSX"))
        excel_file.assert_called_once_with(Path("list.XLSX"), engine="openpyxl")

    try:
        open_excel_workbook(Path("list.xlsm"))
    except ValueError as exc:
        assert str(exc) == "Chỉ hỗ trợ file Excel .xls và .xlsx."
    else:
        raise AssertionError("Unsupported Excel extensions should be rejected")


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
                    ([(0, 0), (10, 0), (10, 5), (0, 5)], ("CCCD 034202005650", 0.8)),
                    ([(20, 10), (30, 10), (30, 15), (20, 15)], ("23/10/2002", 0.6)),
                    ([], (" ", 1.0)),
                ]

        with pdfplumber.open(pdf_path) as pdf, patch("app.get_deepdoc_ocr", return_value=FakeOCR()):
            rendered = np.array(pdf.pages[0].to_image(resolution=216).annotated)
            assert tuple(rendered.shape) == (
                round(page.rect.height * 3),
                round(page.rect.width * 3),
                3,
            )
            text, confidence, lines, layout = ocr_page(pdf.pages[0])

        assert text == "CCCD 034202005650 23/10/2002"
        assert abs(confidence - 72.5925925926) < 0.001
        assert lines == [
            ("CCCD 034202005650", 80.0),
            ("23/10/2002", 60.0),
        ]
        assert layout[0][:2] == ("CCCD 034202005650", 80.0)
        assert layout[0][2] == (
            0.0, 0.0, 10 / rendered.shape[1], 5 / rendered.shape[0]
        )


def test_ocr_page_rotates_image_and_maps_boxes_back_to_page_coordinates():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        page = document.new_page()
        document.save(pdf_path)

        class FakeOCR:
            def __call__(self, image):
                assert image.shape[0] == round(page.rect.width * 3)
                assert image.shape[1] == round(page.rect.height * 3)
                return [
                    ([(0, 0), (10, 0), (10, 5), (0, 5)], ("DOB", 0.9)),
                ]

        with pdfplumber.open(pdf_path) as pdf, patch(
            "app.get_deepdoc_ocr", return_value=FakeOCR()
        ):
            _, _, _, layout = ocr_page(pdf.pages[0], rotation_degrees=90)

        assert layout[0][2] == (
            0.0,
            1 - 10 / round(pdf.pages[0].height * 3),
            5 / round(pdf.pages[0].width * 3),
            1.0,
        )


def test_ocr_birth_date_parser_prefers_slashes_and_uses_id_year_when_malformed():
    assert extract_ocr_birth_dates("DOB 29/02/1.888") == ["29/02/1888"]
    assert extract_ocr_birth_dates(
        "DOB 29/02/1.888", inferred_birth_year=1988
    ) == ["29/02/1988"]
    assert extract_ocr_birth_dates(
        "DOB 29/02/1.888; other date 01.03.1988",
        inferred_birth_year=1988,
    ) == ["29/02/1988"]
    assert extract_ocr_birth_dates(
        "DOB 13/1111991", inferred_birth_year=1991
    ) == ["13/11/1991"]
    assert extract_ocr_birth_dates(
        "Ngày tinh/Lan dinh 11.07/2003"
    ) == ["11/07/2003"]


def test_spatial_birth_date_vote_parses_huy_mixed_separator_reading():
    text = "Ngày tinh/Lan dinh 11.07/2003"
    pages = [[(text, 61.5)]]
    layout_pages = [[(text, 61.5, (0.1, 0.1, 0.6, 0.15))]]

    assert extract_birth_date_from_pages(
        pages,
        today=date(2026, 11, 13),
        layout_pages=layout_pages,
    ) == ("11/07/2003", 61.5, False, False, False)


def test_vietnamese_id_birth_year_inference_requires_valid_12_digit_id():
    assert vietnamese_id_birth_year("033088001705") == 1988
    assert vietnamese_id_birth_year("001202005650") == 2002
    assert vietnamese_id_birth_year("145892756") is None


def test_inspect_pdf_infers_uncertain_dob_year_only_from_high_confidence_pdf_id():
    for id_confidence, expected_dates in ((85.0, {"29/02/1988"}), (79.9, set())):
        with tempfile.TemporaryDirectory() as td:
            pdf_path = Path(td) / "scan.pdf"
            document = Document()
            document.new_page()
            document.save(pdf_path)
            lines = [
                ("CCCD 033088001705", id_confidence),
                ("Ngày sinh: 29/02/1.888", 90.0),
            ]
            layout = [
                ("CCCD 033088001705", id_confidence, (0.1, 0.1, 0.4, 0.15)),
                ("Ngày sinh: 29/02/1.888", 90.0, (0.1, 0.2, 0.6, 0.25)),
            ]
            with patch(
                "app.ocr_page",
                return_value=(
                    " ".join(text for text, _ in lines),
                    88.0,
                    lines,
                    layout,
                ),
            ):
                info = inspect_pdf(pdf_path)

            assert info.birth_dates == expected_dates


def test_inspect_pdf_retries_rotations_and_keeps_best_dob_evidence():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        label = ("Date of birth:", 90.0, (0.1, 0.1, 0.3, 0.15))

        def fake_ocr_page(page, rotation_degrees=0):
            if rotation_degrees == 0:
                return (
                    "Date of birth unreadable",
                    80.0,
                    [("Date of birth:", 90.0), ("unreadable", 60.0)],
                    [label],
                )
            if rotation_degrees == 90:
                return (
                    "23/06/1983",
                    90.0,
                    [("23/06/1983", 90.0)],
                    [("23/06/1983", 90.0, (0.1, 0.16, 0.3, 0.2))],
                )
            return (
                "20/06/1983",
                60.0,
                [("20/06/1983", 60.0)],
                [("20/06/1983", 60.0, (0.1, 0.16, 0.3, 0.2))],
            )

        with patch("app.ocr_page", side_effect=fake_ocr_page) as ocr_mock:
            info = inspect_pdf(pdf_path)

        assert info.birth_dates == {"23/06/1983"}
        assert [call.kwargs.get("rotation_degrees", 0) for call in ocr_mock.call_args_list] == [
            0, 90, -90
        ]


def test_inspect_pdf_retries_when_ocr_confidence_is_below_65_percent():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        unrelated = ("Unrelated text", 64.9, (0.1, 0.1, 0.3, 0.15))
        dob = ("Date of birth: 23/06/1983", 90.0, (0.1, 0.1, 0.6, 0.15))

        def fake_ocr_page(page, rotation_degrees=0):
            if rotation_degrees == 0:
                return "Unrelated text", 64.9, [("Unrelated text", 64.9)], [unrelated]
            if rotation_degrees == 90:
                return (
                    dob[0],
                    90.0,
                    [(dob[0], 90.0)],
                    [dob],
                )
            return "Unrelated text", 60.0, [("Unrelated text", 60.0)], [unrelated]

        with patch("app.ocr_page", side_effect=fake_ocr_page) as ocr_mock:
            info = inspect_pdf(pdf_path)

        assert info.birth_dates == {"23/06/1983"}
        assert [
            call.kwargs.get("rotation_degrees", 0)
            for call in ocr_mock.call_args_list
        ] == [0, 90, -90]


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


def test_spatial_dob_votes_prefer_repeated_labeled_date_over_higher_single_score():
    pages = [
        [
            ("Date of birth", 90.0),
            ("23/06/1983", 91.1),
            ("Ngày sinh:", 88.9),
            ("23/06/1983", 91.8),
        ],
        [
            ("23/06/1983", 72.0),
            ("Ngày tháng năm sinh:", 89.2),
        ],
        [
            ("Ngày sinh / Date of binh 20/06/1983", 75.9),
        ],
    ]
    layout_pages = [
        [
            ("Date of birth", 90.0, (0.05, 0.1, 0.25, 0.15)),
            ("23/06/1983", 91.1, (0.35, 0.1, 0.55, 0.15)),
            ("Ngày sinh:", 88.9, (0.05, 0.5, 0.25, 0.55)),
            ("23/06/1983", 91.8, (0.35, 0.5, 0.55, 0.55)),
        ],
        [
            ("23/06/1983", 72.0, (0.35, 0.4, 0.55, 0.45)),
            ("Ngày tháng năm sinh:", 89.2, (0.05, 0.4, 0.25, 0.45)),
        ],
        [
            ("Ngày sinh / Date of binh 20/06/1983", 75.9, (0.37, 0.40, 0.66, 0.43)),
        ],
    ]

    assert extract_birth_date_from_pages(
        pages,
        today=date(2026, 11, 13),
        layout_pages=layout_pages,
    ) == ("23/06/1983", (91.1 + 91.8 + 72.0) / 3, False, False, False)


def test_spatial_dob_vote_tie_requires_manual_review():
    pages = [
        [("Date of birth:", 90.0), ("23/06/1983", 90.0)],
        [("Date of birth:", 90.0), ("20/06/1983", 90.0)],
    ]
    layout_pages = [
        [
            ("Date of birth:", 90.0, (0.1, 0.1, 0.3, 0.15)),
            ("23/06/1983", 90.0, (0.1, 0.16, 0.3, 0.2)),
        ],
        [
            ("Date of birth:", 90.0, (0.1, 0.1, 0.3, 0.15)),
            ("20/06/1983", 90.0, (0.1, 0.16, 0.3, 0.2)),
        ],
    ]

    assert extract_birth_date_from_pages(
        pages,
        today=date(2026, 11, 13),
        layout_pages=layout_pages,
    ) == (None, None, False, True, False)


def test_spatial_dob_distance_uses_smaller_axis_gap_for_same_row():
    pages = [[("Date of birth:", 90.0), ("23/06/1983", 90.0)]]
    layout_pages = [[
        ("Date of birth:", 90.0, (0.1, 0.1, 0.2, 0.15)),
        ("23/06/1983", 90.0, (0.35, 0.1, 0.45, 0.15)),
    ]]

    assert extract_birth_date_from_pages(
        pages,
        today=date(2026, 11, 13),
        layout_pages=layout_pages,
    ) == ("23/06/1983", 90.0, False, False, False)


def test_spatial_dob_distance_rejects_candidate_separated_on_both_axes():
    pages = [[("Date of birth:", 90.0), ("23/06/1983", 90.0)]]
    layout_pages = [[
        ("Date of birth:", 90.0, (0.1, 0.1, 0.2, 0.15)),
        ("23/06/1983", 90.0, (0.35, 0.22, 0.45, 0.27)),
    ]]

    assert extract_birth_date_from_pages(
        pages,
        today=date(2026, 11, 13),
        layout_pages=layout_pages,
    ) == (None, None, False, False, False)


def test_spatial_dob_vote_margin_threshold_is_point_zero_zero_four():
    pages = [
        [("Date of birth:", 90.0), ("23/06/1983", 90.0)],
        [("Date of birth:", 90.0), ("20/06/1983", 74.0)],
    ]
    layout_pages = [
        [
            ("Date of birth:", 90.0, (0.1, 0.1, 0.3, 0.15)),
            ("23/06/1983", 90.0, (0.1, 0.16, 0.3, 0.2)),
        ],
        [
            ("Date of birth:", 90.0, (0.1, 0.1, 0.3, 0.15)),
            ("20/06/1983", 74.0, (0.1, 0.16, 0.3, 0.2)),
        ],
    ]

    assert extract_birth_date_from_pages(
        pages,
        today=date(2026, 11, 13),
        layout_pages=layout_pages,
    ) == ("23/06/1983", 90.0, False, False, False)


def test_spatial_dob_vote_margin_at_point_zero_zero_four_is_ambiguous():
    pages = [
        [("Date of birth:", 90.0), ("23/06/1983", 90.0)],
        [("Date of birth:", 90.0), ("20/06/1983", 89.2)],
    ]
    layout_pages = [
        [
            ("Date of birth:", 90.0, (0.1, 0.1, 0.3, 0.15)),
            ("23/06/1983", 90.0, (0.1, 0.16, 0.3, 0.2)),
        ],
        [
            ("Date of birth:", 90.0, (0.1, 0.1, 0.3, 0.15)),
            ("20/06/1983", 89.2, (0.1, 0.16, 0.3, 0.2)),
        ],
    ]

    assert extract_birth_date_from_pages(
        pages,
        today=date(2026, 11, 13),
        layout_pages=layout_pages,
    ) == (None, None, False, True, False)


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
            return_value=(" ".join(text for text, _ in lines), 85.0, lines, []),
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
            return_value=("CCCD: 034202 005650", 70.0, lines, []),
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
            return_value=(" ".join(text for text, _ in lines), 85.0, lines, []),
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
            return_value=(" ".join(text for text, _ in lines), 85.0, lines, []),
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
            return_value=("recognized text", 86.7, [("recognized text", 86.7)], []),
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
                [],
            ),
        ):
            info = inspect_pdf(pdf_path)

        assert info.ids == {"034202005650"}
        assert info.id_confidence == 90.0
        assert info.id_from_pdf


def test_inspect_pdf_recovers_valid_id_substring_from_overlong_ocr_digits():
    cases = [
        ("001187008952", "1001187008952"),
        ("033183009701", "10033183009701"),
        ("001302000802", "1001302000802"),
        ("033083008215", "10033083008215"),
        ("033079009842", "10033079009842"),
        ("033097001450", "10033097001450"),
        ("001200037461", "061001200037461"),
    ]

    for expected_id, ocr_line in cases:
        with tempfile.TemporaryDirectory() as td:
            pdf_path = Path(td) / f"scan {expected_id}.pdf"
            document = Document()
            document.new_page()
            document.save(pdf_path)
            with patch(
                "app.ocr_page",
                return_value=(" ".join(["recognized"] * 8 + [ocr_line]), 75.0,
                              [("recognized", 75.0)] * 8 + [(ocr_line, 75.0)], []),
            ):
                info = inspect_pdf(pdf_path)

            assert info.ids == {expected_id}
            assert info.id_confidence == 75.0
            assert info.id_from_pdf


def test_overlong_ocr_id_recovery_does_not_accept_invalid_window():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan 027074009701.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        with patch(
            "app.ocr_page",
            return_value=(
                "recognized text",
                75.0,
                [("recognized text", 75.0), ("110427074009701", 75.0)],
                [],
            ),
        ):
            info = inspect_pdf(pdf_path)

        assert info.ids == {"027074009701"}
        assert not info.id_from_pdf


def test_inspect_pdf_uses_filename_9_digit_id_when_pdf_has_no_12_digit_id():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "Bui Quang Tiep 145892756.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        with patch(
            "app.ocr_page",
            return_value=("unreadable", 80.0, [("unreadable", 80.0)], []),
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


def test_match_document_prefers_filename_id_match_with_closer_filename_name():
    worse_name_match = DocumentInfo(
        Path("Nguyen Huong Vy 030073005631.pdf"),
        "", set(), set(), set(), "nguyen huong vy", 80.0, True,
        filename_ids={"030073005631"},
    )
    better_name_match = DocumentInfo(
        Path("Nguyen Luong Huy 030073005631.pdf"),
        "", set(), set(), set(), "nguyen luong huy", 80.0, True,
        filename_ids={"030073005631"},
    )

    assert match_document(
        "Nguyễn Lương Huy",
        "030073005631",
        [worse_name_match, better_name_match],
    ) == (better_name_match, 1.0)


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
    test_excel_reader_engine_is_selected_by_extension()
    test_ocr_page_reads_deepdoc_result_text_and_score()
    test_ocr_page_rotates_image_and_maps_boxes_back_to_page_coordinates()
    test_ocr_birth_date_parser_prefers_slashes_and_uses_id_year_when_malformed()
    test_spatial_birth_date_vote_parses_huy_mixed_separator_reading()
    test_vietnamese_id_birth_year_inference_requires_valid_12_digit_id()
    test_inspect_pdf_infers_uncertain_dob_year_only_from_high_confidence_pdf_id()
    test_inspect_pdf_retries_rotations_and_keeps_best_dob_evidence()
    test_inspect_pdf_retries_when_ocr_confidence_is_below_65_percent()
    test_extract_birth_date_from_line_before_label()
    test_extract_birth_date_when_ocr_label_has_extra_noise()
    test_extract_birth_date_ignores_low_confidence_label_or_date()
    test_extract_birth_date_ignores_date_outside_nearby_lines()
    test_nearby_birth_date_beats_higher_confidence_distant_date()
    test_wide_birth_date_fallback_is_flagged_for_manual_review()
    test_nearby_birth_date_on_another_page_beats_wide_fallback()
    test_birth_date_search_does_not_cross_pdf_pages()
    test_spatial_dob_votes_prefer_repeated_labeled_date_over_higher_single_score()
    test_spatial_dob_vote_tie_requires_manual_review()
    test_spatial_dob_distance_uses_smaller_axis_gap_for_same_row()
    test_spatial_dob_distance_rejects_candidate_separated_on_both_axes()
    test_spatial_dob_vote_margin_threshold_is_point_zero_zero_four()
    test_spatial_dob_vote_margin_at_point_zero_zero_four_is_ambiguous()
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
    test_inspect_pdf_recovers_valid_id_substring_from_overlong_ocr_digits()
    test_overlong_ocr_id_recovery_does_not_accept_invalid_window()
    test_inspect_pdf_uses_filename_9_digit_id_when_pdf_has_no_12_digit_id()
    test_match_document_does_not_use_filename_name_without_pdf_id()
    test_match_document_prefers_filename_id_match_with_closer_filename_name()
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
