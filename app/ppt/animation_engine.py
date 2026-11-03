from __future__ import annotations

import copy
from dataclasses import dataclass
from xml.etree import ElementTree as ET

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}

ET.register_namespace("p", P_NS)
ET.register_namespace("a", A_NS)


@dataclass(slots=True)
class AnimationConfig:
    template_shape_name: str = "LYRIC_TEMPLATE"
    effect_in: str = "fade"
    effect_out: str = "fade"


def apply_template_animation(slide, ordered_shape_names: list[str], config: AnimationConfig | None = None) -> None:
    cfg = config or AnimationConfig()
    if len(ordered_shape_names) < 2:
        return

    xml = slide._element.xml
    root = ET.fromstring(xml.encode("utf-8"))

    shape_id_by_name = _shape_id_lookup(root)
    if cfg.template_shape_name not in shape_id_by_name:
        return
    if ordered_shape_names[0] != cfg.template_shape_name:
        return

    timing = root.find(".//p:timing", NS)
    if timing is None:
        return

    main_seq = timing.find(".//p:seq", NS)
    if main_seq is None:
        return

    template_effects = _find_effect_nodes_for_shape(main_seq, shape_id_by_name[cfg.template_shape_name])
    if not template_effects:
        return

    for shape_name in ordered_shape_names[1:]:
        shape_id = shape_id_by_name.get(shape_name)
        if shape_id is None:
            continue
        for template_effect in template_effects:
            cloned = copy.deepcopy(template_effect)
            _retarget_effect_shape_id(cloned, shape_id)
            main_seq.append(cloned)

    slide._element.clear()
    updated = ET.tostring(root, encoding="unicode")
    rebuilt = ET.fromstring(updated.encode("utf-8"))
    slide._element.extend(list(rebuilt))
    slide._element.attrib.update(rebuilt.attrib)


def _shape_id_lookup(root: ET.Element) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for sp in root.findall(".//p:sp", NS):
        c_nv_pr = sp.find("./p:nvSpPr/p:cNvPr", NS)
        if c_nv_pr is None:
            continue
        name = c_nv_pr.attrib.get("name")
        shape_id = c_nv_pr.attrib.get("id")
        if name and shape_id:
            mapping[name] = shape_id
    return mapping


def _find_effect_nodes_for_shape(main_seq: ET.Element, shape_id: str) -> list[ET.Element]:
    effects: list[ET.Element] = []
    for node in list(main_seq):
        tgt = node.find(".//p:spTgt", NS)
        if tgt is not None and tgt.attrib.get("spid") == shape_id:
            effects.append(node)
    return effects


def _retarget_effect_shape_id(effect_node: ET.Element, new_shape_id: str) -> None:
    target = effect_node.find(".//p:spTgt", NS)
    if target is not None:
        target.attrib["spid"] = new_shape_id
