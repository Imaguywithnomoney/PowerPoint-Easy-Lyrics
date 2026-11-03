from __future__ import annotations

from pathlib import Path
from typing import Protocol

from app.ai.schema import LyricStanza


class _AnimationConfig(Protocol):
    animation_in: str
    animation_out: str

PP_LAYOUT_BLANK = 12
PP_SAVE_AS_OPEN_XML_PRESENTATION = 24
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
PP_ALIGN_CENTER = 2
MSO_ANCHOR_MIDDLE = 3
MSO_ANIM_TRIGGER_ON_CLICK = 1
MSO_ANIM_TRIGGER_WITH_PREVIOUS = 2
MSO_ANIM_EFFECT_APPEAR = 1
MSO_ANIM_EFFECT_FADE = 10
MSO_ANIM_EFFECT_DISSOLVE = 14
MSO_ANIM_EFFECT_WIPE = 22

SLIDE_WIDTH_POINTS = 960.0
SLIDE_HEIGHT_POINTS = 540.0


def export_with_powerpoint_com(grouped_stanzas: list[list[LyricStanza]], output_path: Path, config: _AnimationConfig) -> None:
    import win32com.client  # type: ignore[import-not-found]

    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    ppt.Visible = True
    ppt.DisplayAlerts = 0
    presentation = ppt.Presentations.Add()
    saved = False
    try:
        presentation.PageSetup.SlideWidth = SLIDE_WIDTH_POINTS
        presentation.PageSetup.SlideHeight = SLIDE_HEIGHT_POINTS
        for group in grouped_stanzas:
            _add_animated_slide(presentation, group, config)
        presentation.SaveAs(str(output_path.resolve()), PP_SAVE_AS_OPEN_XML_PRESENTATION)
        saved = True
    finally:
        try:
            presentation.Close()
        except Exception:
            if not saved:
                raise
        try:
            ppt.Quit()
        except Exception:
            if not saved:
                raise


def _add_animated_slide(presentation, stanzas: list[LyricStanza], config: _AnimationConfig) -> None:
    slide = presentation.Slides.Add(presentation.Slides.Count + 1, PP_LAYOUT_BLANK)
    slide.FollowMasterBackground = False
    slide.Background.Fill.ForeColor.RGB = _rgb_to_office((0, 0, 0))

    left = 72.0
    top = 90.0
    width = 816.0
    height = 360.0

    shapes = []
    for stanza in stanzas:
        shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height)
        text_range = shape.TextFrame.TextRange
        text_range.Text = stanza.text
        text_range.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        text_range.Font.Name = "Arial"
        text_range.Font.Size = 48
        text_range.Font.Bold = True
        text_range.Font.Color.RGB = _rgb_to_office((255, 255, 255))
        shape.TextFrame2.VerticalAnchor = MSO_ANCHOR_MIDDLE
        shape.TextFrame2.WordWrap = True
        shapes.append(shape)

    _apply_click_sequence(slide, shapes, config.animation_in, config.animation_out)


def _apply_click_sequence(slide, shapes, effect_in: str, effect_out: str) -> None:
    if len(shapes) <= 1:
        return

    main_seq = slide.TimeLine.MainSequence
    in_effect = _entrance_effect(effect_in)
    out_effect = _exit_effect(effect_out)

    # Entrance on later stanzas hides them until clicked; first stanza stays visible.
    for index in range(len(shapes) - 1):
        exit_effect = main_seq.AddEffect(shapes[index], out_effect, 0, MSO_ANIM_TRIGGER_ON_CLICK)
        exit_effect.Exit = True
        main_seq.AddEffect(shapes[index + 1], in_effect, 0, MSO_ANIM_TRIGGER_WITH_PREVIOUS)


def _entrance_effect(name: str) -> int:
    effect = name.lower().strip()
    if effect in {"appear", "disappear"}:
        return MSO_ANIM_EFFECT_APPEAR
    if effect == "wipe":
        return MSO_ANIM_EFFECT_WIPE
    if effect == "dissolve":
        return MSO_ANIM_EFFECT_DISSOLVE
    return MSO_ANIM_EFFECT_FADE


def _exit_effect(name: str) -> int:
    effect = name.lower().strip()
    if effect in {"appear", "disappear"}:
        return MSO_ANIM_EFFECT_APPEAR
    if effect == "wipe":
        return MSO_ANIM_EFFECT_WIPE
    if effect == "dissolve":
        return MSO_ANIM_EFFECT_DISSOLVE
    return MSO_ANIM_EFFECT_FADE


def _rgb_to_office(rgb: tuple[int, int, int]) -> int:
    r, g, b = rgb
    return (b << 16) | (g << 8) | r
