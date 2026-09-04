from __future__ import annotations

from pathlib import Path

from pptx.util import Inches


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
ASSET_DIR = ROOT / "assets"
OUTPUT_DIR = ROOT / "output"

ASSUMPTIONS_PATH = DATA_DIR / "assumptions.yaml"
SOURCES_PATH = DATA_DIR / "public_sources.yaml"
SLIDE_CONTENT_PATH = DATA_DIR / "slide_content.yaml"

PPTX_PATH = OUTPUT_DIR / "replit_series_e_investor_deck.pptx"
XLSX_PATH = OUTPUT_DIR / "replit_series_e_sourcebook.xlsx"

SLIDE_WIDTH = Inches(13.333333)
SLIDE_HEIGHT = Inches(7.5)

COLORS = {
    "orange": "FF3C00",
    "orange_bright": "FF3C00",
    "blue": "2492FF",
    "coral": "FFB199",
    "cream": "FAF6F1",
    "paper": "FFFDFC",
    "ink": "181818",
    "charcoal": "312E2E",
    "gray_700": "767270",
    "gray_500": "A8A5A3",
    "gray_300": "C5C5C5",
    "gray_200": "DBD4CF",
    "gray_100": "F5F3F0",
    "white": "FFFFFF",
    "green": "228B5A",
    "red": "C33B32",
    "yellow": "E8A317",
}

FONT_SANS = "Arial"
FONT_MONO = "Courier New"

EVIDENCE_LABELS = {
    "public": "Public fact",
    "derived": "Derived",
    "estimate": "Est.",
    "management_target": "Management target",
}
