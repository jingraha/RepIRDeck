from __future__ import annotations

from openpyxl import load_workbook
from pptx import Presentation

from src.build import build_all
from src.xlsx_builder import SHEET_NAMES


def test_build_outputs(tmp_path) -> None:
    pptx_path, xlsx_path = build_all(tmp_path)
    presentation = Presentation(pptx_path)
    assert len(presentation.slides) == 30
    slide_text = {
        index + 1: "\n".join(
            shape.text
            for shape in slide.shapes
            if getattr(shape, "has_text_frame", False)
        )
        for index, slide in enumerate(presentation.slides)
    }
    assert "$20.8B" in slide_text[2]
    assert "73x" in slide_text[13]
    assert "52%" in slide_text[24]
    assert "27% FCF margin" in slide_text[27]

    workbook = load_workbook(xlsx_path, read_only=True, data_only=False)
    assert workbook.sheetnames == SHEET_NAMES
    assert workbook["QA Checks"]["D5"].value == "PASS"
    workbook.close()
