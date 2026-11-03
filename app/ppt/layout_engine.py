from __future__ import annotations

from dataclasses import dataclass

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from pptx.util import Pt

from app.ai.schema import LyricStanza


@dataclass(slots=True)
class WorshipTheme:
    background_rgb: tuple[int, int, int] = (0, 0, 0)
    text_rgb: tuple[int, int, int] = (255, 255, 255)
    font_name: str = "Arial"
    font_size_pt: int = 50
    margin_left_in: float = 1.0
    margin_top_in: float = 1.6
    box_width_in: float = 11.333
    box_height_in: float = 3.5


@dataclass(slots=True)
class TextBoxSpec:
    left: int
    top: int
    width: int
    height: int


def build_presentation_base() -> Presentation:
    presentation = Presentation()
    # 16:9 widescreen output.
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    return presentation


def apply_worship_background(slide, theme: WorshipTheme) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*theme.background_rgb)


def compute_textbox_spec(theme: WorshipTheme) -> TextBoxSpec:
    emu_per_in = 914400
    return TextBoxSpec(
        left=int(theme.margin_left_in * emu_per_in),
        top=int(theme.margin_top_in * emu_per_in),
        width=int(theme.box_width_in * emu_per_in),
        height=int(theme.box_height_in * emu_per_in),
    )


def add_stacked_stanza_shapes(
    slide,
    stanzas: list[LyricStanza],
    theme: WorshipTheme,
    template_shape_name: str = "LYRIC_TEMPLATE",
) -> list[str]:
    shape_names: list[str] = []
    spec = compute_textbox_spec(theme)

    for index, stanza in enumerate(stanzas):
        shape = slide.shapes.add_textbox(spec.left, spec.top, spec.width, spec.height)
        shape.name = template_shape_name if index == 0 else f"LYRIC_{index + 1}"

        text_frame = shape.text_frame
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        text_frame.text = stanza.text
        for paragraph in text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(theme.font_size_pt)
                run.font.bold = True
                run.font.name = theme.font_name
                run.font.color.rgb = RGBColor(*theme.text_rgb)

        shape_names.append(shape.name)
    return shape_names
