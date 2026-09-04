from __future__ import annotations

from src.sources import load_slide_content, load_sources


def test_every_slide_has_sources() -> None:
    sources = load_sources()
    content = load_slide_content()
    assert set(content["source_map"]) == set(range(1, 31))
    for slide, source_ids in content["source_map"].items():
        assert source_ids, f"Slide {slide} has no sources"
        assert all(source_id in sources for source_id in source_ids)


def test_public_sources_have_urls() -> None:
    sources = load_sources()
    for source in sources.values():
        if source.evidence_class == "public":
            assert source.url
