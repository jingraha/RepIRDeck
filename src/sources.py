from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .config import SLIDE_CONTENT_PATH, SOURCES_PATH
from .model import load_yaml


@dataclass(frozen=True)
class Source:
    id: str
    publisher: str
    title: str
    url: str | None
    publication_date: str | None
    access_date: str
    quality: str
    evidence_class: str
    claims: tuple[str, ...]

    @classmethod
    def from_mapping(cls, payload: dict[str, Any]) -> "Source":
        return cls(
            id=str(payload["id"]),
            publisher=str(payload["publisher"]),
            title=str(payload["title"]),
            url=payload.get("url"),
            publication_date=payload.get("publication_date"),
            access_date=str(payload["access_date"]),
            quality=str(payload["quality"]),
            evidence_class=str(payload["evidence_class"]),
            claims=tuple(payload.get("claims", [])),
        )


def load_sources(path: Path = SOURCES_PATH) -> dict[str, Source]:
    payload = load_yaml(path)
    sources = [Source.from_mapping(item) for item in payload["sources"]]
    result = {source.id: source for source in sources}
    if len(result) != len(sources):
        raise ValueError("Source IDs must be unique")
    return result


def load_slide_content(path: Path = SLIDE_CONTENT_PATH) -> dict[str, Any]:
    return load_yaml(path)


def format_slide_sources(
    slide_number: int,
    sources: dict[str, Source],
    source_map: dict[int, list[str]],
) -> str:
    source_ids = source_map.get(slide_number, [])
    return "Sources: " + " | ".join(sources[source_id].id for source_id in source_ids)


def source_rows(sources: dict[str, Source]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for source in sources.values():
        if source.claims:
            for claim in source.claims:
                rows.append(
                    {
                        "source_id": source.id,
                        "publisher": source.publisher,
                        "title": source.title,
                        "claim": claim,
                        "publication_date": source.publication_date,
                        "access_date": source.access_date,
                        "quality": source.quality,
                        "evidence_class": source.evidence_class,
                        "url": source.url,
                    }
                )
        else:
            rows.append(
                {
                    "source_id": source.id,
                    "publisher": source.publisher,
                    "title": source.title,
                    "claim": "",
                    "publication_date": source.publication_date,
                    "access_date": source.access_date,
                    "quality": source.quality,
                    "evidence_class": source.evidence_class,
                    "url": source.url,
                }
            )
    return rows
