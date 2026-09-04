from __future__ import annotations

import argparse
from pathlib import Path

from .config import PPTX_PATH, XLSX_PATH
from .model import build_model
from .pptx_builder import build_presentation
from .sources import load_slide_content, load_sources
from .validation import validate_inputs, validate_outputs
from .xlsx_builder import SHEET_NAMES, build_workbook


def build_all(output_dir: Path | None = None) -> tuple[Path, Path]:
    model = build_model()
    sources = load_sources()
    slide_content = load_slide_content()
    validate_inputs(model, sources, slide_content)

    pptx_path = (output_dir / PPTX_PATH.name) if output_dir else PPTX_PATH
    xlsx_path = (output_dir / XLSX_PATH.name) if output_dir else XLSX_PATH
    build_workbook(model, sources, xlsx_path)
    build_presentation(model, sources, slide_content, pptx_path)
    validate_outputs(pptx_path, xlsx_path, SHEET_NAMES)
    return pptx_path, xlsx_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build the Replit Series E investor deck and sourcebook."
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Override the default output directory.",
    )
    args = parser.parse_args()
    pptx_path, xlsx_path = build_all(args.output_dir)
    print(f"Built {pptx_path}")
    print(f"Built {xlsx_path}")


if __name__ == "__main__":
    main()
