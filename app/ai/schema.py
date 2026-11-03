from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class SectionLabel(str, Enum):
    VERSE = "verse"
    CHORUS = "chorus"
    BRIDGE = "bridge"
    TAG = "tag"
    REPEAT = "repeat"
    UNKNOWN = "unknown"


@dataclass(slots=True)
class LyricStanza:
    text: str
    label: SectionLabel = SectionLabel.UNKNOWN
    group_hint: str | None = None

    def validate(self) -> None:
        if not self.text or not self.text.strip():
            raise ValueError("Stanza text cannot be empty.")


@dataclass(slots=True)
class SegmentationResult:
    title: str | None = None
    stanzas: list[LyricStanza] = field(default_factory=list)

    def validate(self) -> None:
        if not self.stanzas:
            raise ValueError("At least one stanza is required.")
        for stanza in self.stanzas:
            stanza.validate()


def result_to_json_payload(result: SegmentationResult) -> dict[str, Any]:
    result.validate()
    return {
        "title": result.title,
        "stanzas": [
            {
                "text": stanza.text,
                "label": stanza.label.value,
                "group_hint": stanza.group_hint,
            }
            for stanza in result.stanzas
        ],
    }


def parse_segmentation_payload(payload: dict[str, Any]) -> SegmentationResult:
    if "stanzas" not in payload or not isinstance(payload["stanzas"], list):
        raise ValueError("Invalid payload: missing 'stanzas' list.")

    stanzas: list[LyricStanza] = []
    for row in payload["stanzas"]:
        if not isinstance(row, dict):
            raise ValueError("Invalid payload: each stanza must be an object.")
        label_raw = str(row.get("label", "unknown")).lower().strip()
        label = SectionLabel(label_raw) if label_raw in SectionLabel._value2member_map_ else SectionLabel.UNKNOWN
        stanzas.append(
            LyricStanza(
                text=str(row.get("text", "")).strip(),
                label=label,
                group_hint=(str(row.get("group_hint")).strip() if row.get("group_hint") else None),
            )
        )

    result = SegmentationResult(title=(str(payload.get("title")).strip() if payload.get("title") else None), stanzas=stanzas)
    result.validate()
    return result


OLLAMA_JSON_PROMPT = """You are a lyrics segmenter for church worship presentation software.
Return valid JSON only, no markdown.
Schema:
{
  "title": "optional string",
  "stanzas": [
    {
      "text": "stanza text with line breaks preserved",
      "label": "verse|chorus|bridge|tag|repeat|unknown",
      "group_hint": "optional short hint"
    }
  ]
}
Rules:
1) Split lyrics into natural stanzas.
2) Preserve original wording and line breaks.
3) Guess section labels conservatively.
4) Never return empty stanzas.
"""
