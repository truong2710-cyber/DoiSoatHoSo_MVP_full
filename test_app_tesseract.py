import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from fitz import Document
import pdfplumber

from app_tesseract import ocr_page, parse_tesseract_tsv


TSV_HEADER = (
    "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\t"
    "width\theight\tconf\ttext\n"
)


def test_tesseract_tsv_groups_words_into_lines_and_averages_confidence():
    tsv = TSV_HEADER + (
        "5\t1\t1\t1\t1\t1\t20\t20\t50\t10\t90\tNgày\n"
        "5\t1\t1\t1\t1\t2\t75\t18\t35\t12\t80\tsinh:\n"
        "5\t1\t1\t1\t2\t1\t20\t40\t100\t12\t70\t05/10/1993\n"
    )

    text, confidence, lines, layout = parse_tesseract_tsv(
        tsv, image_width=200, image_height=100
    )

    assert text == "Ngày sinh: 05/10/1993"
    assert lines == [
        ("Ngày sinh:", (4 * 90 + 5 * 80) / 9),
        ("05/10/1993", 70.0),
    ]
    assert confidence == (
        len("Ngày sinh:") * lines[0][1] + len("05/10/1993") * 70
    ) / (len("Ngày sinh:") + len("05/10/1993"))
    assert layout == [
        ("Ngày sinh:", lines[0][1], (0.1, 0.18, 0.55, 0.3)),
        ("05/10/1993", 70.0, (0.1, 0.4, 0.6, 0.52)),
    ]


def test_tesseract_tsv_maps_rotated_boxes_to_original_page_coordinates():
    tsv = TSV_HEADER + (
        "5\t1\t1\t1\t1\t1\t10\t20\t30\t10\t90\tDOB\n"
    )

    _, _, _, layout = parse_tesseract_tsv(
        tsv, image_width=200, image_height=100, rotation_degrees=90
    )

    assert layout[0][2] == (0.2, 0.8, 0.3, 0.95)


def test_ocr_page_uses_tesseract_tsv_backend():
    with tempfile.TemporaryDirectory() as td:
        pdf_path = Path(td) / "scan.pdf"
        document = Document()
        document.new_page()
        document.save(pdf_path)
        tsv = TSV_HEADER + (
            "5\t1\t1\t1\t1\t1\t10\t20\t30\t10\t90\tDOB\n"
        )
        with patch(
            "app_tesseract.get_tesseract_config",
            return_value=("/usr/bin/tesseract", "vie+eng"),
        ), patch(
            "app_tesseract.subprocess.run",
            return_value=SimpleNamespace(
                returncode=0,
                stdout=tsv.encode("utf-8"),
                stderr=b"",
            ),
        ) as run, pdfplumber.open(pdf_path) as pdf:
            text, confidence, lines, layout = ocr_page(pdf.pages[0])

        assert text == "DOB"
        assert confidence == 90.0
        assert lines == [("DOB", 90.0)]
        assert layout[0][0:2] == ("DOB", 90.0)
        assert run.call_args.args[0][1:5] == [
            "stdin", "stdout", "-l", "vie+eng"
        ]
        assert run.call_args.kwargs["input"].startswith(b"\x89PNG")


if __name__ == "__main__":
    test_tesseract_tsv_groups_words_into_lines_and_averages_confidence()
    test_tesseract_tsv_maps_rotated_boxes_to_original_page_coordinates()
    test_ocr_page_uses_tesseract_tsv_backend()
    print("All Tesseract tests passed")
