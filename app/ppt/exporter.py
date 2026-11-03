from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.ai.segmenter import SegmenterConfig, segment_lyrics, split_for_slides
from app.ppt.animation_engine import AnimationConfig, apply_template_animation
from app.ppt.layout_engine import (
    WorshipTheme,
    add_stacked_stanza_shapes,
    apply_worship_background,
    build_presentation_base,
)


@dataclass(slots=True)
class ExportConfig:
    max_stanzas_per_slide: int = 4
    use_local_ai: bool = False
    ollama_model: str = "llama3.1:8b-instruct-q4_K_M"
    animation_in: str = "fade"
    animation_out: str = "fade"
    max_lines_per_stanza: int = 4
    use_powerpoint_com: bool = True
    require_powerpoint_com: bool = False


def export_lyrics_to_pptx(lyrics: str, output_path: str | Path, config: ExportConfig | None = None) -> Path:
    cfg = config or ExportConfig()
    segment_result = segment_lyrics(
        lyrics,
        SegmenterConfig(
            model=cfg.ollama_model,
            use_ai=cfg.use_local_ai,
            max_lines_per_stanza=cfg.max_lines_per_stanza,
        ),
    )
    grouped = split_for_slides(segment_result.stanzas, cfg.max_stanzas_per_slide)
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)

    if cfg.use_powerpoint_com:
        try:
            from app.ppt.com_backend import export_with_powerpoint_com

            export_with_powerpoint_com(grouped, destination, cfg)
            return destination
        except Exception:
            if cfg.require_powerpoint_com:
                raise

    presentation = build_presentation_base()
    blank_layout = presentation.slide_layouts[6]

    for group in grouped:
        slide = presentation.slides.add_slide(blank_layout)
        apply_worship_background(slide, WorshipTheme())
        shape_names = add_stacked_stanza_shapes(slide, group, WorshipTheme())
        apply_template_animation(
            slide,
            shape_names,
            AnimationConfig(effect_in=cfg.animation_in, effect_out=cfg.animation_out),
        )

    presentation.save(destination.as_posix())
    return destination
