from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Iterable
from urllib.error import URLError
from urllib.request import Request, urlopen

from app.ai.schema import (
    OLLAMA_JSON_PROMPT,
    LyricStanza,
    SectionLabel,
    SegmentationResult,
    parse_segmentation_payload,
)


@dataclass(slots=True)
class SegmenterConfig:
    model: str = "llama3.1:8b-instruct-q4_K_M"
    endpoint: str = "http://127.0.0.1:11434/api/generate"
    timeout_seconds: int = 45
    use_ai: bool = True
    max_lines_per_stanza: int = 4


SECTION_HINTS: dict[SectionLabel, tuple[str, ...]] = {
    SectionLabel.CHORUS: ("chorus", "refrain"),
    SectionLabel.VERSE: ("verse",),
    SectionLabel.BRIDGE: ("bridge",),
    SectionLabel.TAG: ("tag", "ending"),
    SectionLabel.REPEAT: ("repeat",),
}


def segment_lyrics(raw_lyrics: str, config: SegmenterConfig | None = None) -> SegmentationResult:
    cfg = config or SegmenterConfig()
    normalized = raw_lyrics.replace("\r\n", "\n").strip()
    if not normalized:
        raise ValueError("Lyrics cannot be empty.")

    if cfg.use_ai:
        ai_result = _segment_with_ollama(normalized, cfg)
        if ai_result is not None:
            return _enforce_stanza_line_limit(ai_result, cfg.max_lines_per_stanza)
    return _enforce_stanza_line_limit(_fallback_segment(normalized), cfg.max_lines_per_stanza)


def _segment_with_ollama(text: str, config: SegmenterConfig) -> SegmentationResult | None:
    payload = {
        "model": config.model,
        "prompt": f"{OLLAMA_JSON_PROMPT}\n\nLyrics:\n{text}",
        "stream": False,
        "format": "json",
    }
    request = Request(
        config.endpoint,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=config.timeout_seconds) as response:
            body = json.loads(response.read().decode("utf-8"))
    except (URLError, TimeoutError, OSError, json.JSONDecodeError):
        return None

    model_output = body.get("response", "")
    if not model_output:
        return None

    try:
        parsed = json.loads(model_output)
        return parse_segmentation_payload(parsed)
    except (json.JSONDecodeError, ValueError):
        return None


def _fallback_segment(text: str) -> SegmentationResult:
    chunks = [chunk.strip() for chunk in re.split(r"\n\s*\n+", text) if chunk.strip()]
    if len(chunks) <= 1:
        chunks = _line_based_chunks(text, max_lines=4)
    stanzas: list[LyricStanza] = []
    seen_lines: set[str] = set()

    for chunk in chunks:
        label = _guess_label(chunk)
        normalized_line = _normalize_repeat_key(chunk)
        if normalized_line in seen_lines:
            label = SectionLabel.REPEAT
        seen_lines.add(normalized_line)
        stanzas.append(LyricStanza(text=chunk, label=label))

    result = SegmentationResult(title=None, stanzas=stanzas)
    result.validate()
    return result


def _line_based_chunks(text: str, max_lines: int) -> list[str]:
    raw_lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not raw_lines:
        return []
    chunks: list[str] = []
    current: list[str] = []
    for line in raw_lines:
        current.append(line)
        if len(current) >= max_lines:
            chunks.append("\n".join(current))
            current = []
    if current:
        chunks.append("\n".join(current))
    return chunks


def _enforce_stanza_line_limit(result: SegmentationResult, max_lines_per_stanza: int) -> SegmentationResult:
    if max_lines_per_stanza < 1:
        return result
    rewritten: list[LyricStanza] = []
    for stanza in result.stanzas:
        lines = [line.strip() for line in stanza.text.splitlines() if line.strip()]
        if len(lines) <= max_lines_per_stanza:
            rewritten.append(stanza)
            continue
        for idx in range(0, len(lines), max_lines_per_stanza):
            rewritten.append(
                LyricStanza(
                    text="\n".join(lines[idx : idx + max_lines_per_stanza]),
                    label=stanza.label,
                    group_hint=stanza.group_hint,
                )
            )
    constrained = SegmentationResult(title=result.title, stanzas=rewritten)
    constrained.validate()
    return constrained


def _guess_label(stanza_text: str) -> SectionLabel:
    first_line = stanza_text.splitlines()[0].strip().lower()
    for label, hints in SECTION_HINTS.items():
        if any(hint in first_line for hint in hints):
            return label
    return SectionLabel.UNKNOWN


def _normalize_repeat_key(stanza_text: str) -> str:
    return " ".join(
        line.strip().lower() for line in stanza_text.splitlines() if line.strip()
    )


def split_for_slides(stanzas: Iterable[LyricStanza], max_stanzas_per_slide: int) -> list[list[LyricStanza]]:
    if max_stanzas_per_slide < 1:
        raise ValueError("max_stanzas_per_slide must be >= 1")
    grouped: list[list[LyricStanza]] = []
    current: list[LyricStanza] = []
    for stanza in stanzas:
        current.append(stanza)
        if len(current) == max_stanzas_per_slide:
            grouped.append(current)
            current = []
    if current:
        grouped.append(current)
    return grouped
