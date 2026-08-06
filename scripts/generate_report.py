"""Generate a 广发策略 DOCX by replacing slots in the retained template.

This generator deliberately edits the OOXML package rather than creating a
new ``Document``.  Headers, footers, sections, styles, theme parts, custom XML
and legal pages remain template-owned.  Native tables, images and charts are
copied from the source report as editable Word objects.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import date
import itertools
import json
import math
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import shutil
import tempfile
import unicodedata
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from asset_copy import (
    PackagePartImporter,
    clone_with_imported_relationships,
)
from chart_formatter import apply_chart_template


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = ROOT_DIR / "template" / "总量通用.docx"
DEFAULT_STRUCTURE = ROOT_DIR / "runtime" / "clean_structure.json"
DEFAULT_MANIFEST = ROOT_DIR / "runtime" / "assets_manifest.json"
DEFAULT_LAYOUT = ROOT_DIR / "config" / "template_layout.json"
DEFAULT_CHART_TEMPLATE = ROOT_DIR / "template" / "图表模板案例新版.xlsx"
DEFAULT_OUTPUT_DIR = ROOT_DIR / "output"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
NS = {
    "w": W_NS, "r": R_NS, "a": A_NS, "c": C_NS,
    "wp": WP_NS, "pic": PIC_NS, "pr": PR_NS,
}


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


W_RPR_ORDER = {
    name: index for index, name in enumerate((
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
        "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
        "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
        "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
        "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
        "eastAsianLayout", "specVanish", "oMath", "rPrChange",
    ))
}
W_PPR_ORDER = {
    name: index for index, name in enumerate((
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
        "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
        "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
        "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
        "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
        "pPrChange",
    ))
}
W_TBLPR_ORDER = {
    name: index for index, name in enumerate((
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
        "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
        "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
        "tblCaption", "tblDescription", "tblPrChange",
    ))
}
W_TCPR_ORDER = {
    name: index for index, name in enumerate((
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
        "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
        "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
    ))
}


def reorder_word_properties(parent: etree._Element, order: dict[str, int]) -> None:
    """Keep edited property children in the strict OOXML schema sequence."""
    children = list(parent)
    ranked = sorted(
        enumerate(children),
        key=lambda item: (
            order.get(etree.QName(item[1]).localname, len(order)), item[0]
        ),
    )
    parent[:] = [child for _, child in ranked]


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def clean_filename_title(filename: str) -> str:
    name = Path(filename).stem
    for suffix in (
        "_内容版", "内容版", "_初稿", "初稿", "_v1", "_V1",
        "_v2", "_V2", "_final", "_FINAL",
    ):
        name = name.replace(suffix, "")
    return name.strip()


def output_path_for(filename: str, output_dir: Path) -> Path:
    title = clean_filename_title(filename).replace("/", "_").replace("\\", "_")
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / f"【广发策略】{title}.docx"


def element_text(element: etree._Element) -> str:
    return "".join(element.xpath(".//w:t/text()", namespaces=NS)).strip()


def clear_paragraph(paragraph: etree._Element) -> None:
    for child in list(paragraph):
        if child.tag != qn(W_NS, "pPr"):
            paragraph.remove(child)


def set_paragraph_text(paragraph: etree._Element, text: str) -> None:
    run_properties = paragraph.find(".//w:r/w:rPr", namespaces=NS)
    clear_paragraph(paragraph)
    if not text:
        return
    run = etree.SubElement(paragraph, qn(W_NS, "r"))
    if run_properties is not None:
        run.append(deepcopy(run_properties))
    text_node = etree.SubElement(run, qn(W_NS, "t"))
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text


def ensure_run_properties(run: etree._Element) -> etree._Element:
    properties = run.find("w:rPr", namespaces=NS)
    if properties is None:
        properties = etree.Element(qn(W_NS, "rPr"))
        run.insert(0, properties)
    return properties


def set_run_font(
    run: etree._Element,
    font_name: str,
    size_half_points: int | None = None,
) -> None:
    properties = ensure_run_properties(run)
    fonts = properties.find("w:rFonts", namespaces=NS)
    if fonts is None:
        fonts = etree.Element(qn(W_NS, "rFonts"))
        properties.insert(0, fonts)
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(W_NS, attribute), font_name)
    for theme_attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(W_NS, theme_attribute), None)
    if size_half_points is not None:
        for local in ("sz", "szCs"):
            node = properties.find(f"w:{local}", namespaces=NS)
            if node is None:
                node = etree.SubElement(properties, qn(W_NS, local))
            node.set(qn(W_NS, "val"), str(size_half_points))
    for local in ("b", "bCs"):
        node = properties.find(f"w:{local}", namespaces=NS)
        if node is None:
            node = etree.SubElement(properties, qn(W_NS, local))
        node.set(qn(W_NS, "val"), "0")
    reorder_word_properties(properties, W_RPR_ORDER)


def paragraph_from_prototype(
    prototype: etree._Element,
    text: str,
    runs: list[dict] | None = None,
    text_config: dict | None = None,
) -> etree._Element:
    paragraph = deepcopy(prototype)
    if not text_config:
        set_paragraph_text(paragraph, text)
        return paragraph
    clear_paragraph(paragraph)
    light_font = text_config.get("normal_font", "思源黑体 CN Light")
    medium_font = text_config.get("emphasis_font", "思源黑体 CN Medium")
    usable_runs = runs if runs and "".join(r.get("text", "") for r in runs) == text else None
    for segment in usable_runs or [{"text": text, "emphasis": False}]:
        value = segment.get("text", "")
        if not value:
            continue
        run = append_run(paragraph, value)
        set_run_font(run, medium_font if segment.get("emphasis") else light_font)
    return paragraph


def style_maps(styles_root: etree._Element) -> tuple[dict[str, str], dict[str, str]]:
    name_to_id: dict[str, str] = {}
    id_to_name: dict[str, str] = {}
    for style in styles_root.xpath("./w:style", namespaces=NS):
        style_id = style.get(qn(W_NS, "styleId"))
        name_node = style.find("w:name", namespaces=NS)
        if not style_id or name_node is None:
            continue
        name = name_node.get(qn(W_NS, "val"))
        name_to_id[name] = style_id
        id_to_name[style_id] = name
    return name_to_id, id_to_name


def paragraph_style_id(paragraph: etree._Element) -> str | None:
    node = paragraph.find("w:pPr/w:pStyle", namespaces=NS)
    return node.get(qn(W_NS, "val")) if node is not None else None


def find_prototypes(
    document_root: etree._Element,
    styles_root: etree._Element,
    configured_styles: dict,
) -> dict[str, etree._Element]:
    name_to_id, _ = style_maps(styles_root)
    prototypes: dict[str, etree._Element] = {}
    paragraphs = document_root.xpath("//w:body//w:p", namespaces=NS)
    for role, style_name in configured_styles.items():
        style_id = name_to_id.get(style_name)
        match = next((p for p in paragraphs if paragraph_style_id(p) == style_id), None)
        if match is None:
            raise ValueError(f"Template style has no paragraph prototype: {style_name}")
        prototypes[role] = deepcopy(match)
    return prototypes


def find_marker_paragraph(root: etree._Element, marker: str) -> etree._Element:
    for paragraph in root.xpath(".//w:p", namespaces=NS):
        if element_text(paragraph) == marker:
            return paragraph
    raise ValueError(f"Template marker not found: {marker}")


def row_offset_paragraph(marker_paragraph: etree._Element, offset: int) -> etree._Element:
    marker_row = marker_paragraph.xpath("ancestor::w:tr[1]", namespaces=NS)[0]
    table = marker_row.getparent()
    rows = table.findall("w:tr", namespaces=NS)
    index = rows.index(marker_row) + offset
    target = rows[index].find("w:tc", namespaces=NS)
    paragraph = target.find("w:p", namespaces=NS)
    if paragraph is None:
        paragraph = etree.SubElement(target, qn(W_NS, "p"))
    return paragraph


def clear_template_markers(document_root: etree._Element) -> None:
    """Remove non-printing implementation controls from retained components."""
    for paragraph in document_root.xpath("//w:p", namespaces=NS):
        if re.fullmatch(r"\[Table_[^]]+\]", element_text(paragraph)):
            set_paragraph_text(paragraph, "")
    for text_node in document_root.xpath("//w:t", namespaces=NS):
        if text_node.text:
            text_node.text = re.sub(r"X{2,}", "", text_node.text)


def fill_cover(
    document_root: etree._Element,
    styles_root: etree._Element,
    layout: dict,
    title: str,
    summary: list[dict],
) -> None:
    cover = layout["cover"]
    title_marker = find_marker_paragraph(document_root, cover["title_marker"])
    title_paragraph = row_offset_paragraph(title_marker, cover["title_row_offset"])
    set_paragraph_text(title_paragraph, title)
    subtitle = row_offset_paragraph(title_marker, cover["subtitle_row_offset"])
    set_paragraph_text(subtitle, "")
    set_paragraph_text(title_marker, "")

    summary_marker = find_marker_paragraph(document_root, cover["summary_marker"])
    summary_row = summary_marker.xpath("ancestor::w:tr[1]", namespaces=NS)[0]
    summary_table = summary_row.getparent()
    rows = summary_table.findall("w:tr", namespaces=NS)
    target_row = rows[rows.index(summary_row) + cover["summary_row_offset"]]
    target_cell = target_row.find("w:tc", namespaces=NS)
    old_paragraphs = target_cell.findall("w:p", namespaces=NS)
    prototype = old_paragraphs[0] if old_paragraphs else etree.Element(qn(W_NS, "p"))
    name_to_id, _ = style_maps(styles_root)
    risk_style_id = name_to_id.get(cover.get("risk_style", "首页正文"))
    risk_prototype = next(
        (p for p in old_paragraphs if paragraph_style_id(p) == risk_style_id),
        prototype,
    )
    for child in list(target_cell):
        if child.tag != qn(W_NS, "tcPr"):
            target_cell.remove(child)
    text_config = layout.get("text", {})
    for index, item in enumerate(summary):
        # Every cover-summary paragraph follows the reference's compact
        # 首页正文 rhythm; emphasis is handled at run level.
        source = risk_prototype
        target_cell.append(
            paragraph_from_prototype(
                source,
                item.get("text", ""),
                item.get("runs"),
                text_config,
            )
        )
    if not summary:
        target_cell.append(paragraph_from_prototype(prototype, ""))
    set_paragraph_text(summary_marker, "核心观点")

    industry_marker = find_marker_paragraph(document_root, cover["industry_date_marker"])
    industry_line = row_offset_paragraph(
        industry_marker, cover["industry_date_row_offset"]
    )
    set_paragraph_text(industry_line, "证券研究报告 ｜ 策略周报")
    set_paragraph_text(industry_marker, "")

    top_table = document_root.xpath("//w:body/w:tbl", namespaces=NS)[
        cover["top_level_table_index"]
    ]
    top_rows = top_table.findall("w:tr", namespaces=NS)
    for row_index in cover.get("clear_sidebar_top_level_rows", []):
        if row_index >= len(top_rows):
            continue
        cells = top_rows[row_index].findall("w:tc", namespaces=NS)
        if len(cells) < 3:
            continue
        cell = cells[2]
        cell_properties = cell.find("w:tcPr", namespaces=NS)
        for child in list(cell):
            if child is not cell_properties:
                cell.remove(child)
        cell.append(etree.Element(qn(W_NS, "p")))

    # Remaining symbolic template markers are implementation controls, not
    # report content.  Clear them everywhere while retaining their containers.
    clear_template_markers(document_root)


def split_front_matter(content: list[dict], fallback_title: str) -> tuple[str, list[dict], list[dict]]:
    first_heading = next(
        (i for i, item in enumerate(content) if item.get("type") == "heading1"),
        len(content),
    )
    front = content[:first_heading]
    body = content[first_heading:]
    title = fallback_title
    summary: list[dict] = []
    title_consumed = False
    for item in front:
        if item.get("type") != "paragraph":
            continue
        text = item.get("text", "").strip()
        if not text or text in {"【摘要】", "摘要"}:
            continue
        if not title_consumed:
            title = text
            title_consumed = True
        else:
            summary.append(item)
    return title, summary, body


def field_paragraph(
    instruction: str,
    prototype: etree._Element | None = None,
) -> etree._Element:
    paragraph = deepcopy(prototype) if prototype is not None else etree.Element(qn(W_NS, "p"))
    clear_paragraph(paragraph)
    begin_run = etree.SubElement(paragraph, qn(W_NS, "r"))
    begin = etree.SubElement(begin_run, qn(W_NS, "fldChar"))
    begin.set(qn(W_NS, "fldCharType"), "begin")
    instr_run = etree.SubElement(paragraph, qn(W_NS, "r"))
    instr = etree.SubElement(instr_run, qn(W_NS, "instrText"))
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = f" {instruction} "
    separate_run = etree.SubElement(paragraph, qn(W_NS, "r"))
    separate = etree.SubElement(separate_run, qn(W_NS, "fldChar"))
    separate.set(qn(W_NS, "fldCharType"), "separate")
    end_run = etree.SubElement(paragraph, qn(W_NS, "r"))
    end = etree.SubElement(end_run, qn(W_NS, "fldChar"))
    end.set(qn(W_NS, "fldCharType"), "end")
    return paragraph


def cached_field_paragraphs(
    instruction: str,
    entries: list[tuple[int, str]],
    level_prototypes: dict[int, etree._Element],
    field_prototype: etree._Element,
    closing_prototype: etree._Element,
) -> list[etree._Element]:
    """Build a native field with deterministic visible cached entries.

    Word TOC fields normally contain their cached result between the outer
    ``separate`` and ``end`` markers.  Keeping all three markers in one empty
    paragraph makes the index blank until Word elects to refresh it.  The
    generated cache below remains visible in headless/public-platform output,
    while the surrounding field can still be refreshed by Word later.
    """

    visible_entries = entries or [(1, "暂无")]
    result: list[etree._Element] = []
    for index, (level, text) in enumerate(visible_entries):
        prototype = level_prototypes.get(level)
        if prototype is None:
            prototype = level_prototypes.get(1)
        if prototype is None:
            prototype = field_prototype
        paragraph = deepcopy(prototype)
        clear_paragraph(paragraph)
        if index == 0:
            begin_run = etree.SubElement(paragraph, qn(W_NS, "r"))
            begin = etree.SubElement(begin_run, qn(W_NS, "fldChar"))
            begin.set(qn(W_NS, "fldCharType"), "begin")
            begin.set(qn(W_NS, "dirty"), "true")
            instruction_run = etree.SubElement(paragraph, qn(W_NS, "r"))
            instruction_node = etree.SubElement(
                instruction_run, qn(W_NS, "instrText")
            )
            instruction_node.set(
                "{http://www.w3.org/XML/1998/namespace}space", "preserve"
            )
            instruction_node.text = f" {instruction} "
            separator_run = etree.SubElement(paragraph, qn(W_NS, "r"))
            separator = etree.SubElement(separator_run, qn(W_NS, "fldChar"))
            separator.set(qn(W_NS, "fldCharType"), "separate")
        append_run(paragraph, text)
        result.append(paragraph)

    closing = deepcopy(closing_prototype)
    clear_paragraph(closing)
    end_run = etree.SubElement(closing, qn(W_NS, "r"))
    end = etree.SubElement(end_run, qn(W_NS, "fldChar"))
    end.set(qn(W_NS, "fldCharType"), "end")
    result.append(closing)
    return result


def chinese_number(value: int) -> str:
    digits = "零一二三四五六七八九"
    if value < 10:
        return digits[value]
    if value < 20:
        return "十" + (digits[value % 10] if value % 10 else "")
    if value < 100:
        return digits[value // 10] + "十" + (digits[value % 10] if value % 10 else "")
    return str(value)


def chinese_number_value(value: str) -> int | None:
    digits = {character: index for index, character in enumerate("零一二三四五六七八九")}
    if value in digits:
        return digits[value]
    if "十" not in value:
        return None
    left, right = value.split("十", 1)
    tens = digits.get(left, 1) if left else 1
    units = digits.get(right, 0) if right else 0
    return tens * 10 + units


def risk_heading_for_content(body_content: list[dict]) -> str:
    ordinals: list[int] = []
    for item in body_content:
        if item.get("type") != "heading1":
            continue
        match = re.match(r"^([一二三四五六七八九十]+)、", item.get("text", "").strip())
        if not match:
            continue
        ordinal = chinese_number_value(match.group(1))
        if ordinal is not None:
            ordinals.append(ordinal)
    next_ordinal = max(ordinals, default=0) + 1
    return f"{chinese_number(next_ordinal)}、风险提示"


def index_entries(
    body_content: list[dict],
    risk_heading: str,
    chart_caption_fallbacks: dict[str, str] | None = None,
) -> tuple[list[tuple[int, str]], list[tuple[int, str]], list[tuple[int, str]]]:
    main: list[tuple[int, str]] = []
    figures: list[tuple[int, str]] = []
    tables: list[tuple[int, str]] = []
    counts = {"图": 0, "表": 0}
    chart_caption_fallbacks = chart_caption_fallbacks or {}

    def add_caption(raw_text: str, kind: str) -> None:
        text = raw_text.strip()
        if not text:
            return
        counts[kind] += 1
        title = re.sub(r"^[图表]\s*\d*\s*[：:]?\s*", "", text)
        target = tables if kind == "表" else figures
        target.append((1, f"{kind} {counts[kind]}：{title}" if title else f"{kind} {counts[kind]}"))

    pending_caption: tuple[str, str] | None = None
    for item in body_content:
        item_type = item.get("type")
        if item_type == "heading1":
            main.append((1, item.get("text", "")))
        elif item_type == "heading2":
            main.append((2, item.get("text", "")))
        elif item_type in {"figure_caption", "table_caption"}:
            kind = "表" if item_type == "table_caption" else "图"
            pending_caption = (item.get("text", ""), kind)
        elif item_type == "figure":
            raw_text = item.get("caption", "") or chart_caption_fallbacks.get(
                item.get("asset_id", ""), ""
            )
            kind = "表" if raw_text.lstrip().startswith("表") else "图"
            if pending_caption is not None:
                raw_text, kind = pending_caption
            add_caption(raw_text, kind)
            pending_caption = None
        elif item_type == "table":
            visual_groups = item.get("visual_groups", [])
            if visual_groups:
                for group in visual_groups:
                    raw_text = group.get("caption", "")
                    kind = "表" if raw_text.lstrip(" '▪").startswith("表") else "图"
                    add_caption(raw_text, kind)
                pending_caption = None
            elif pending_caption is not None:
                raw_text, kind = pending_caption
                add_caption(raw_text, "表" if kind == "表" else kind)
                pending_caption = None
    main.append((1, risk_heading))
    return main, figures, tables


def page_break_paragraph() -> etree._Element:
    paragraph = etree.Element(qn(W_NS, "p"))
    run = etree.SubElement(paragraph, qn(W_NS, "r"))
    br = etree.SubElement(run, qn(W_NS, "br"))
    br.set(qn(W_NS, "type"), "page")
    return paragraph


def body_section_breaks(body: etree._Element) -> list[etree._Element]:
    return [
        child for child in body
        if child.tag == qn(W_NS, "p") and child.find("w:pPr/w:sectPr", namespaces=NS) is not None
    ]


def replace_index(
    body: etree._Element,
    body_start: etree._Element,
    layout: dict,
    body_content: list[dict],
    risk_heading: str,
    chart_caption_fallbacks: dict[str, str] | None = None,
) -> None:
    section_breaks = body_section_breaks(body)
    ordinal = layout["index"]["replace_after_section_break"]
    if ordinal < 1 or ordinal > len(section_breaks):
        raise ValueError("Invalid index section break locator")
    first_break = section_breaks[ordinal - 1]
    children = list(body)
    start = children.index(first_break) + 1
    end = children.index(body_start)
    old_index = children[start:end]

    def paragraph_with_text(value: str) -> etree._Element:
        match = next(
            (
                child for child in old_index
                if child.tag == qn(W_NS, "p") and element_text(child) == value
            ),
            None,
        )
        if match is None:
            raise ValueError(f"Template index prototype not found: {value}")
        return deepcopy(match)

    def field_prototype(pattern: str) -> etree._Element:
        normalized_pattern = re.sub(r"\s+", "", pattern)
        match = next(
            (
                child for child in old_index
                if child.tag == qn(W_NS, "p")
                and normalized_pattern in re.sub(
                    r"\s+",
                    "",
                    " ".join(child.xpath(".//w:instrText/text()", namespaces=NS)),
                )
            ),
            None,
        )
        if match is None:
            raise ValueError(f"Template field prototype not found: {pattern}")
        return deepcopy(match)

    blank_prototype = next(
        (
            deepcopy(child) for child in old_index
            if child.tag == qn(W_NS, "p") and not element_text(child)
            and not child.xpath(".//w:instrText", namespaces=NS)
            and not child.xpath(".//w:br", namespaces=NS)
        ),
        etree.Element(qn(W_NS, "p")),
    )
    break_prototype = next(
        (
            deepcopy(child) for child in old_index
            if child.tag == qn(W_NS, "p")
            and child.xpath(".//w:br[@w:type='page']", namespaces=NS)
        ),
        page_break_paragraph(),
    )
    main_title = paragraph_with_text(layout["index"]["title"])
    visual_title = paragraph_with_text(layout["index"]["visual_title"])
    main_prototype = field_prototype('TOC \\o')
    figure_prototype = field_prototype('TOC \\h \\z \\c "图"')
    table_prototype = field_prototype('TOC \\h \\z \\c "表格"')

    main_style = paragraph_style_id(main_prototype)
    main_level2_prototype = next(
        (
            deepcopy(child) for child in old_index
            if child.tag == qn(W_NS, "p")
            and element_text(child).lstrip().startswith(("（", "("))
            and paragraph_style_id(child) != main_style
        ),
        deepcopy(main_prototype),
    )
    closing_prototype = next(
        (
            deepcopy(child) for child in old_index
            if child.tag == qn(W_NS, "p")
            and child.xpath(".//w:fldChar[@w:fldCharType='end']", namespaces=NS)
            and not child.xpath(".//w:instrText", namespaces=NS)
            and not element_text(child)
        ),
        deepcopy(blank_prototype),
    )

    for child in children[start:end]:
        body.remove(child)
    insertion = body.index(body_start)
    toc_levels = layout["index"].get("toc_levels", "1-2")
    figure_sequence = layout["index"].get("figure_sequence", "图")
    table_sequence = layout["index"].get("table_sequence", "表")
    main_entries, figure_entries, table_entries = index_entries(
        body_content, risk_heading, chart_caption_fallbacks
    )
    replacements = [
        main_title,
        deepcopy(blank_prototype),
    ]
    replacements.extend(
        cached_field_paragraphs(
            f'TOC \\o "{toc_levels}" \\h \\z \\u',
            main_entries,
            {1: main_prototype, 2: main_level2_prototype},
            main_prototype,
            closing_prototype,
        )
    )
    if layout["index"].get("insert_page_break_between_indexes", True):
        replacements.append(deepcopy(break_prototype))
    replacements.extend(
        [
            visual_title,
            deepcopy(blank_prototype),
        ]
    )
    replacements.extend(
        cached_field_paragraphs(
            f'TOC \\h \\z \\c "{figure_sequence}"',
            figure_entries,
            {1: figure_prototype},
            figure_prototype,
            closing_prototype,
        )
    )
    replacements.append(deepcopy(blank_prototype))
    replacements.extend(
        cached_field_paragraphs(
            f'TOC \\h \\z \\c "{table_sequence}"',
            table_entries,
            {1: table_prototype},
            table_prototype,
            closing_prototype,
        )
    )
    if layout["index"].get("insert_page_break_after", True):
        replacements.append(deepcopy(break_prototype))
    for offset, replacement in enumerate(replacements):
        body.insert(insertion + offset, replacement)


def source_visual_paragraph(
    source_block: etree._Element,
    object_type: str,
    object_index: int,
) -> etree._Element:
    xpath = ".//c:chart" if object_type == "chart" else ".//a:blip"
    objects = source_block.xpath(xpath, namespaces=NS)
    if object_index >= len(objects):
        raise ValueError(f"Visual object index out of range: {object_type}[{object_index}]")
    drawing = objects[object_index].xpath("ancestor::w:drawing[1]", namespaces=NS)
    if not drawing:
        raise ValueError(f"Unsupported visual container for {object_type}")
    paragraph = etree.Element(qn(W_NS, "p"))
    source_ppr = source_block.find("w:pPr", namespaces=NS)
    if source_ppr is not None:
        paragraph.append(deepcopy(source_ppr))
    run = etree.SubElement(paragraph, qn(W_NS, "r"))
    run.append(deepcopy(drawing[0]))
    return paragraph


def ensure_paragraph_properties(paragraph: etree._Element) -> etree._Element:
    properties = paragraph.find("w:pPr", namespaces=NS)
    if properties is None:
        properties = etree.Element(qn(W_NS, "pPr"))
        paragraph.insert(0, properties)
    return properties


def set_paragraph_alignment(paragraph: etree._Element, alignment: str) -> None:
    properties = ensure_paragraph_properties(paragraph)
    node = properties.find("w:jc", namespaces=NS)
    if node is None:
        node = etree.SubElement(properties, qn(W_NS, "jc"))
    node.set(qn(W_NS, "val"), alignment)
    reorder_word_properties(properties, W_PPR_ORDER)


def set_keep_next(paragraph: etree._Element) -> None:
    properties = ensure_paragraph_properties(paragraph)
    if properties.find("w:keepNext", namespaces=NS) is None:
        etree.SubElement(properties, qn(W_NS, "keepNext"))
    reorder_word_properties(properties, W_PPR_ORDER)


def set_paragraph_side_indent(
    paragraph: etree._Element,
    side: str,
    value_dxa: int,
) -> None:
    """Set one physical paragraph indent without disturbing other formatting."""
    properties = ensure_paragraph_properties(paragraph)
    indent = properties.find("w:ind", namespaces=NS)
    if indent is None:
        indent = etree.SubElement(properties, qn(W_NS, "ind"))
    indent.set(qn(W_NS, side), str(value_dxa))
    reorder_word_properties(properties, W_PPR_ORDER)


def append_run(
    paragraph: etree._Element,
    text: str | None = None,
    run_properties: etree._Element | None = None,
) -> etree._Element:
    run = etree.SubElement(paragraph, qn(W_NS, "r"))
    if run_properties is not None:
        run.append(deepcopy(run_properties))
    if text is not None:
        node = etree.SubElement(run, qn(W_NS, "t"))
        if text[:1].isspace() or text[-1:].isspace():
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        node.text = text
    return run


def append_sequence_field(
    paragraph: etree._Element,
    sequence_name: str,
    display_number: int,
    run_properties: etree._Element | None,
) -> None:
    begin_run = append_run(paragraph, run_properties=run_properties)
    begin = etree.SubElement(begin_run, qn(W_NS, "fldChar"))
    begin.set(qn(W_NS, "fldCharType"), "begin")
    begin.set(qn(W_NS, "dirty"), "true")
    instruction_run = append_run(paragraph, run_properties=run_properties)
    instruction = etree.SubElement(instruction_run, qn(W_NS, "instrText"))
    instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instruction.text = f" SEQ {sequence_name} \\* ARABIC "
    separator_run = append_run(paragraph, run_properties=run_properties)
    separator = etree.SubElement(separator_run, qn(W_NS, "fldChar"))
    separator.set(qn(W_NS, "fldCharType"), "separate")
    append_run(paragraph, str(display_number), run_properties)
    end_run = append_run(paragraph, run_properties=run_properties)
    end = etree.SubElement(end_run, qn(W_NS, "fldChar"))
    end.set(qn(W_NS, "fldCharType"), "end")


def numbered_caption_paragraph(
    prototype: etree._Element,
    raw_text: str,
    sequence_name: str,
    display_number: int,
    separator_config: dict | None = None,
) -> etree._Element:
    paragraph = deepcopy(prototype)
    run_properties = paragraph.find(".//w:r/w:rPr", namespaces=NS)
    clear_paragraph(paragraph)
    config = separator_config or {}
    caption_font = config.get("font", "思源黑体 CN Medium")
    caption_size = int(config.get("font_size_half_points", 16))
    prefix = config.get("prefix", "'")

    def caption_run(value: str | None = None) -> etree._Element:
        run = append_run(paragraph, value, run_properties)
        set_run_font(run, caption_font, caption_size)
        return run

    title = re.sub(r"^[图表]\s*\d*\s*[：:]?\s*", "", raw_text.strip())
    caption_run(f"{prefix}{sequence_name} ")
    field_start = len(paragraph)
    append_sequence_field(paragraph, sequence_name, display_number, run_properties)
    for run in paragraph[field_start:]:
        if run.tag == qn(W_NS, "r"):
            set_run_font(run, caption_font, caption_size)
    if title:
        caption_run(f"：{title}")
    properties = ensure_paragraph_properties(paragraph)
    for child_name in ("w:keepNext", "w:keepLines", "w:numPr"):
        child = properties.find(child_name, namespaces=NS)
        if child is not None:
            properties.remove(child)
    # The retained caption style inherits a list marker.  Caption formatting
    # is explicit below, so remove the style link as well as direct numbering.
    paragraph_style = properties.find("w:pStyle", namespaces=NS)
    if paragraph_style is not None:
        properties.remove(paragraph_style)
    borders = properties.find("w:pBdr", namespaces=NS)
    if borders is None:
        borders = etree.SubElement(properties, qn(W_NS, "pBdr"))
    bottom = borders.find("w:bottom", namespaces=NS)
    if bottom is None:
        bottom = etree.SubElement(borders, qn(W_NS, "bottom"))
    bottom.set(qn(W_NS, "val"), config.get("style", "single"))
    bottom.set(qn(W_NS, "sz"), str(config.get("size_eighth_points", 4)))
    bottom.set(qn(W_NS, "space"), str(config.get("space_points", 1)))
    bottom.set(qn(W_NS, "color"), config.get("color", "000000"))
    spacing = properties.find("w:spacing", namespaces=NS)
    if spacing is None:
        spacing = etree.SubElement(properties, qn(W_NS, "spacing"))
    spacing.set(qn(W_NS, "after"), str(config.get("after_twips", 40)))
    spacing.set(qn(W_NS, "line"), str(config.get("line_twips", 240)))
    spacing.set(qn(W_NS, "lineRule"), config.get("line_rule", "exact"))
    reorder_word_properties(properties, W_PPR_ORDER)
    return paragraph


def data_source_paragraph(
    layout: dict,
    source_text: str | None = None,
) -> etree._Element:
    """Build the standard source line shown below every figure and table."""
    config = layout.get("data_source", {})
    paragraph = etree.Element(qn(W_NS, "p"))
    properties = etree.SubElement(paragraph, qn(W_NS, "pPr"))
    spacing = etree.SubElement(properties, qn(W_NS, "spacing"))
    spacing.set(qn(W_NS, "before"), str(config.get("before_twips", 20)))
    spacing.set(qn(W_NS, "after"), str(config.get("after_twips", 80)))
    spacing.set(qn(W_NS, "line"), str(config.get("line_twips", 160)))
    spacing.set(qn(W_NS, "lineRule"), config.get("line_rule", "exact"))
    run = append_run(
        paragraph,
        source_text
        or config.get("text", "数据来源：Bloomberg、广发证券发展研究中心"),
    )
    set_run_font(
        run,
        config.get("font", "思源黑体 CN Light"),
        int(config.get("font_size_half_points", 12)),
    )
    return paragraph


def scale_visual_to_width(paragraph: etree._Element, max_width_emu: int) -> None:
    extents = paragraph.xpath(".//wp:extent", namespaces=NS)
    if not extents:
        return
    original_width = int(extents[0].get("cx", "0"))
    original_height = int(extents[0].get("cy", "0"))
    if original_width <= 0 or original_width <= max_width_emu:
        return
    ratio = max_width_emu / original_width
    new_width = max_width_emu
    new_height = max(1, round(original_height * ratio))
    for extent in extents:
        extent.set("cx", str(new_width))
        extent.set("cy", str(new_height))
    for extent in paragraph.xpath(".//a:xfrm/a:ext", namespaces=NS):
        width = int(extent.get("cx", "0"))
        height = int(extent.get("cy", "0"))
        if width > 0:
            extent.set("cx", str(round(width * ratio)))
        if height > 0:
            extent.set("cy", str(round(height * ratio)))


def remove_picture_outline(block: etree._Element) -> None:
    """Remove editable Word picture-frame outlines without altering pixels."""
    for properties in block.xpath(".//pic:spPr", namespaces=NS):
        line = properties.find("a:ln", namespaces=NS)
        if line is None:
            line = etree.SubElement(properties, qn(A_NS, "ln"))
        for child in list(line):
            if etree.QName(child).localname in {
                "noFill", "solidFill", "gradFill", "pattFill", "blipFill",
            }:
                line.remove(child)
        line.insert(0, etree.Element(qn(A_NS, "noFill")))
        drawing_order = {
            name: index for index, name in enumerate((
                "xfrm", "prstGeom", "custGeom", "noFill", "solidFill",
                "gradFill", "blipFill", "pattFill", "grpFill", "ln",
                "effectLst", "effectDag", "scene3d", "sp3d", "extLst",
            ))
        }
        children = list(properties)
        properties[:] = sorted(
            children,
            key=lambda child: drawing_order.get(
                etree.QName(child).localname, len(drawing_order)
            ),
        )


def table_cell(
    width_dxa: int,
    left_margin_dxa: int = 0,
    right_margin_dxa: int = 0,
    vertical_margin_dxa: int = 57,
) -> etree._Element:
    cell = etree.Element(qn(W_NS, "tc"))
    properties = etree.SubElement(cell, qn(W_NS, "tcPr"))
    width = etree.SubElement(properties, qn(W_NS, "tcW"))
    width.set(qn(W_NS, "w"), str(width_dxa))
    width.set(qn(W_NS, "type"), "dxa")
    margins = etree.SubElement(properties, qn(W_NS, "tcMar"))
    for side, value in (
        ("top", vertical_margin_dxa),
        ("left", left_margin_dxa),
        ("bottom", vertical_margin_dxa),
        ("right", right_margin_dxa),
    ):
        margin = etree.SubElement(margins, qn(W_NS, side))
        margin.set(qn(W_NS, "w"), str(value))
        margin.set(qn(W_NS, "type"), "dxa")
    reorder_word_properties(properties, W_TCPR_ORDER)
    return cell


def append_block_to_cell(cell: etree._Element, block: etree._Element) -> None:
    cell.append(block)
    # A Word table cell containing a nested table must end in a paragraph.
    if block.tag == qn(W_NS, "tbl"):
        cell.append(etree.Element(qn(W_NS, "p")))


def normalize_native_table(
    table: etree._Element,
    width_dxa: int,
    indent_dxa: int,
    text_config: dict,
    table_style: dict | None = None,
) -> None:
    """Scale a source table to the configured report grid and font weights."""

    properties = table.find("w:tblPr", namespaces=NS)
    if properties is None:
        properties = etree.Element(qn(W_NS, "tblPr"))
        table.insert(0, properties)
    width = properties.find("w:tblW", namespaces=NS)
    if width is None:
        width = etree.SubElement(properties, qn(W_NS, "tblW"))
    width.set(qn(W_NS, "w"), str(width_dxa))
    width.set(qn(W_NS, "type"), "dxa")
    indent = properties.find("w:tblInd", namespaces=NS)
    if indent is None:
        indent = etree.SubElement(properties, qn(W_NS, "tblInd"))
    indent.set(qn(W_NS, "w"), str(indent_dxa))
    indent.set(qn(W_NS, "type"), "dxa")
    layout_node = properties.find("w:tblLayout", namespaces=NS)
    if layout_node is None:
        layout_node = etree.SubElement(properties, qn(W_NS, "tblLayout"))
    layout_node.set(qn(W_NS, "type"), "fixed")
    reorder_word_properties(properties, W_TBLPR_ORDER)

    grid_columns = table.xpath("./w:tblGrid/w:gridCol", namespaces=NS)
    old_grid = [int(node.get(qn(W_NS, "w"), "0")) for node in grid_columns]
    old_total = sum(old_grid)
    if grid_columns and old_total > 0:
        scaled = [max(1, round(value * width_dxa / old_total)) for value in old_grid]
        scaled[-1] += width_dxa - sum(scaled)
        for node, value in zip(grid_columns, scaled):
            node.set(qn(W_NS, "w"), str(value))
        for row in table.findall("w:tr", namespaces=NS):
            column_index = 0
            for cell in row.findall("w:tc", namespaces=NS):
                span = int(cell.xpath("string(w:tcPr/w:gridSpan/@w:val)", namespaces=NS) or 1)
                cell_width = sum(scaled[column_index : column_index + span])
                column_index += span
                cell_width_node = cell.find("w:tcPr/w:tcW", namespaces=NS)
                if cell_width_node is None:
                    cell_properties = cell.find("w:tcPr", namespaces=NS)
                    if cell_properties is None:
                        cell_properties = etree.Element(qn(W_NS, "tcPr"))
                        cell.insert(0, cell_properties)
                    cell_width_node = etree.SubElement(cell_properties, qn(W_NS, "tcW"))
                cell_width_node.set(qn(W_NS, "w"), str(cell_width))
                cell_width_node.set(qn(W_NS, "type"), "dxa")
                reorder_word_properties(cell_width_node.getparent(), W_TCPR_ORDER)

    light_font = text_config.get("normal_font", "思源黑体 CN Light")
    medium_font = text_config.get("emphasis_font", "思源黑体 CN Medium")
    table_font_size = (
        int(table_style.get("font_size_half_points", 12))
        if table_style
        else None
    )
    for row_index, row in enumerate(table.findall("w:tr", namespaces=NS)):
        for run in row.xpath(".//w:r", namespaces=NS):
            bold = run.find("w:rPr/w:b", namespaces=NS)
            emphasized = row_index == 0 or (
                bold is not None
                and bold.get(qn(W_NS, "val"), "1") not in {"0", "false", "off"}
            )
            set_run_font(
                run,
                medium_font if emphasized else light_font,
                table_font_size,
            )

    if not table_style:
        return

    # Apply only the two fill roles specified by the GF color sheet.  All
    # existing fonts, borders, alignment, spacing, and row geometry remain as
    # supplied by the source/template normalization above.
    header_fill = table_style.get("header_fill", "2E3160")
    header_font_color = table_style.get("header_font_color", "FFFFFF")
    first_column_fill = table_style.get("first_column_fill", "F2F2F2")
    rows = table.findall("w:tr", namespaces=NS)
    for row_index, row in enumerate(rows):
        for column_index, cell in enumerate(row.findall("w:tc", namespaces=NS)):
            if row_index != 0 and column_index != 0:
                continue
            cell_properties = cell.find("w:tcPr", namespaces=NS)
            if cell_properties is None:
                cell_properties = etree.Element(qn(W_NS, "tcPr"))
                cell.insert(0, cell_properties)

            shading = cell_properties.find("w:shd", namespaces=NS)
            if shading is None:
                shading = etree.SubElement(cell_properties, qn(W_NS, "shd"))
            shading.set(qn(W_NS, "val"), "clear")
            shading.set(qn(W_NS, "color"), "auto")
            shading.set(
                qn(W_NS, "fill"),
                header_fill if row_index == 0 else first_column_fill,
            )
            reorder_word_properties(cell_properties, W_TCPR_ORDER)
            if row_index == 0:
                color_properties = []
                for paragraph in cell.xpath(".//w:p", namespaces=NS):
                    paragraph_properties = ensure_paragraph_properties(paragraph)
                    paragraph_run_properties = paragraph_properties.find(
                        "w:rPr", namespaces=NS
                    )
                    if paragraph_run_properties is None:
                        paragraph_run_properties = etree.SubElement(
                            paragraph_properties, qn(W_NS, "rPr")
                        )
                    color_properties.append(paragraph_run_properties)
                    reorder_word_properties(paragraph_properties, W_PPR_ORDER)
                color_properties.extend(
                    ensure_run_properties(run)
                    for run in cell.xpath(".//w:r", namespaces=NS)
                )
                for run_properties in color_properties:
                    color = run_properties.find("w:color", namespaces=NS)
                    if color is None:
                        color = etree.SubElement(run_properties, qn(W_NS, "color"))
                    color.set(qn(W_NS, "val"), header_font_color)
                    for theme_attribute in ("themeColor", "themeTint", "themeShade"):
                        color.attrib.pop(qn(W_NS, theme_attribute), None)
                    reorder_word_properties(run_properties, W_RPR_ORDER)


def table_layout_config(table: etree._Element, layout: dict) -> dict:
    """Choose the standard or full-page grid from estimated Word wrapping."""

    standard = dict(layout.get("figure_layout", {}))
    expanded = layout.get("expanded_table_layout", {})
    rules = layout.get("table_layout_rules", {})
    if not expanded:
        return standard

    rows = table.findall("w:tr", namespaces=NS)
    if not rows:
        return standard

    grid_nodes = table.xpath("./w:tblGrid/w:gridCol", namespaces=NS)
    grid_widths = [
        max(1, int(node.get(qn(W_NS, "w"), "1")))
        for node in grid_nodes
    ]
    column_count = len(grid_widths)
    if not column_count:
        column_count = max(
            (
                sum(
                    int(cell.xpath("string(w:tcPr/w:gridSpan/@w:val)", namespaces=NS) or 1)
                    for cell in row.findall("w:tc", namespaces=NS)
                )
                for row in rows
            ),
            default=1,
        )
        grid_widths = [1] * column_count

    standard_width = int(standard.get("cell_width_dxa", 8051))
    total_grid_width = max(1, sum(grid_widths))
    scaled_widths = [
        max(1, round(width * standard_width / total_grid_width))
        for width in grid_widths
    ]
    scaled_widths[-1] += standard_width - sum(scaled_widths)

    font_half_points = int(
        layout.get("table_style", {}).get("font_size_half_points", 12)
    )
    font_points = max(1.0, font_half_points / 2)
    line_height = int(rules.get("estimated_line_height_twips", 160))
    row_padding = int(rules.get("estimated_row_padding_twips", 80))
    body_cell_lines: list[int] = []
    estimated_height = 0

    def text_units(text: str) -> float:
        units = 0.0
        for character in text:
            if character.isspace():
                units += 0.35
            elif unicodedata.east_asian_width(character) in {"W", "F"}:
                units += 1.0
            else:
                units += 0.55
        return units

    for row_index, row in enumerate(rows):
        column_index = 0
        row_lines = 1
        for cell in row.findall("w:tc", namespaces=NS):
            span = int(
                cell.xpath("string(w:tcPr/w:gridSpan/@w:val)", namespaces=NS)
                or 1
            )
            cell_width = sum(scaled_widths[column_index : column_index + span])
            column_index += span
            # Reserve a small amount for the cell margins before estimating
            # how many 6-point Chinese-character units fit on each line.
            usable_points = max(font_points, cell_width / 20 - 8)
            line_capacity = max(1.0, usable_points / font_points)
            paragraphs = cell.findall("w:p", namespaces=NS)
            cell_lines = 0
            for paragraph in paragraphs or [cell]:
                text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
                explicit_lines = max(
                    1,
                    len(paragraph.xpath(".//w:br", namespaces=NS)) + 1,
                )
                cell_lines += max(
                    explicit_lines,
                    math.ceil(text_units(text) / line_capacity),
                )
            cell_lines = max(1, cell_lines)
            row_lines = max(row_lines, cell_lines)
            if row_index > 0:
                body_cell_lines.append(cell_lines)
        estimated_height += row_lines * line_height + row_padding

    max_cell_lines = max(body_cell_lines, default=1)
    multiline_limit = int(rules.get("expanded_multiline_cell_lines", 3))
    multiline_count = sum(
        line_count >= multiline_limit for line_count in body_cell_lines
    )
    multiline_ratio = multiline_count / max(1, len(body_cell_lines))
    use_expanded = any(
        (
            column_count >= int(rules.get("expanded_min_columns", 6)),
            max_cell_lines > int(rules.get("expanded_max_cell_lines", 3)),
            multiline_ratio
            > float(rules.get("expanded_multiline_cell_ratio", 0.2)),
            estimated_height
            > int(rules.get("expanded_max_estimated_height_twips", 9800)),
        )
    )
    if use_expanded:
        standard.update(expanded)
    return standard


def figure_layout_table(
    caption: etree._Element | None,
    visual: etree._Element,
    layout: dict,
    config_override: dict | None = None,
) -> etree._Element:
    """Position an editable visual using the reference report's fixed grid."""
    config = config_override or layout.get("figure_layout", {})
    table_width = int(config.get("width_dxa", 8051))
    table_indent = int(config.get("indent_dxa", 2689))
    cell_width = int(config.get("cell_width_dxa", table_width))

    table = etree.Element(qn(W_NS, "tbl"))
    properties = etree.SubElement(table, qn(W_NS, "tblPr"))
    width = etree.SubElement(properties, qn(W_NS, "tblW"))
    width.set(qn(W_NS, "w"), str(table_width))
    width.set(qn(W_NS, "type"), "dxa")
    indent = etree.SubElement(properties, qn(W_NS, "tblInd"))
    indent.set(qn(W_NS, "w"), str(table_indent))
    indent.set(qn(W_NS, "type"), "dxa")
    fixed = etree.SubElement(properties, qn(W_NS, "tblLayout"))
    fixed.set(qn(W_NS, "type"), "fixed")

    borders = etree.SubElement(properties, qn(W_NS, "tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = etree.SubElement(borders, qn(W_NS, side))
        border.set(qn(W_NS, "val"), "nil")
    reorder_word_properties(properties, W_TBLPR_ORDER)

    grid = etree.SubElement(table, qn(W_NS, "tblGrid"))
    column = etree.SubElement(grid, qn(W_NS, "gridCol"))
    column.set(qn(W_NS, "w"), str(cell_width))

    if config.get("center_visual", True) and visual.tag == qn(W_NS, "p"):
        set_paragraph_alignment(visual, "center")
    row = etree.SubElement(table, qn(W_NS, "tr"))
    if config.get("keep_together", True):
        row_properties = etree.SubElement(row, qn(W_NS, "trPr"))
        etree.SubElement(row_properties, qn(W_NS, "cantSplit"))
    cell = table_cell(
        cell_width,
        int(config.get("cell_margin_left_dxa", 0)),
        int(config.get("cell_margin_right_dxa", 0)),
        int(config.get("cell_margin_vertical_dxa", 57)),
    )
    if caption is not None:
        append_block_to_cell(cell, caption)
    append_block_to_cell(cell, visual)
    append_block_to_cell(cell, data_source_paragraph(layout))
    row.append(cell)
    return table


def double_figure_layout_table(
    captions: list[etree._Element | None],
    visuals: list[etree._Element],
    layout: dict,
) -> etree._Element:
    """Create the reference report's two-column visual arrangement."""
    config = layout.get("double_figure_layout", {})
    table_width = int(config.get("width_dxa", 10296))
    column_width = int(config.get("column_width_dxa", table_width // 2))
    max_width = int(config.get("max_drawing_width_emu", 3140000))

    table = etree.Element(qn(W_NS, "tbl"))
    properties = etree.SubElement(table, qn(W_NS, "tblPr"))
    width = etree.SubElement(properties, qn(W_NS, "tblW"))
    width.set(qn(W_NS, "w"), str(table_width))
    width.set(qn(W_NS, "type"), "dxa")
    fixed = etree.SubElement(properties, qn(W_NS, "tblLayout"))
    fixed.set(qn(W_NS, "type"), "fixed")
    alignment = etree.SubElement(properties, qn(W_NS, "jc"))
    alignment.set(qn(W_NS, "val"), config.get("alignment", "right"))
    borders = etree.SubElement(properties, qn(W_NS, "tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = etree.SubElement(borders, qn(W_NS, side))
        border.set(qn(W_NS, "val"), "nil")
    reorder_word_properties(properties, W_TBLPR_ORDER)
    grid = etree.SubElement(table, qn(W_NS, "tblGrid"))
    for _ in range(2):
        column = etree.SubElement(grid, qn(W_NS, "gridCol"))
        column.set(qn(W_NS, "w"), str(column_width))

    row = etree.SubElement(table, qn(W_NS, "tr"))
    if config.get("keep_together", True):
        row_properties = etree.SubElement(row, qn(W_NS, "trPr"))
        etree.SubElement(row_properties, qn(W_NS, "cantSplit"))
    center_gap = int(config.get("caption_center_gap_dxa", 114))
    half_gap = center_gap // 2
    for column_index, (caption, visual) in enumerate(zip(captions, visuals)):
        if visual.tag == qn(W_NS, "p"):
            scale_visual_to_width(visual, max_width)
        if config.get("center_visual", True) and visual.tag == qn(W_NS, "p"):
            set_paragraph_alignment(visual, "center")
        cell = table_cell(
            column_width,
            int(config.get("cell_margin_left_dxa", 0)),
            int(config.get("cell_margin_right_dxa", 0)),
            int(config.get("cell_margin_vertical_dxa", 57)),
        )
        if caption is not None:
            # Keep the two independent caption rules visually separate without
            # moving the outer left/right edges or the data-source paragraphs.
            set_paragraph_side_indent(
                caption,
                "right" if column_index == 0 else "left",
                half_gap,
            )
            append_block_to_cell(cell, caption)
        append_block_to_cell(cell, visual)
        append_block_to_cell(cell, data_source_paragraph(layout))
        row.append(cell)
    return table


def import_missing_styles(
    cloned_block: etree._Element,
    source_styles: etree._Element,
    target_styles: etree._Element,
) -> None:
    target_by_name_type: dict[tuple[str, str], str] = {}
    target_defaults: dict[str, str] = {}
    for style in target_styles.xpath("./w:style", namespaces=NS):
        style_id = style.get(qn(W_NS, "styleId"))
        style_type = style.get(qn(W_NS, "type"), "")
        name = style.xpath("string(w:name/@w:val)", namespaces=NS)
        if style_id and name:
            target_by_name_type[(style_type, name)] = style_id
        if style_id and style.get(qn(W_NS, "default")) in {"1", "true", "on"}:
            target_defaults[style_type] = style_id

    style_id_map: dict[str, str] = {}
    for style in source_styles.xpath("./w:style", namespaces=NS):
        source_id = style.get(qn(W_NS, "styleId"))
        style_type = style.get(qn(W_NS, "type"), "")
        name = style.xpath("string(w:name/@w:val)", namespaces=NS)
        if not source_id:
            continue
        matching_target = target_by_name_type.get((style_type, name))
        if matching_target:
            style_id_map[source_id] = matching_target
        elif style.get(qn(W_NS, "default")) in {"1", "true", "on"}:
            matching_default = target_defaults.get(style_type)
            if matching_default:
                style_id_map[source_id] = matching_default

    for node in cloned_block.xpath(
        ".//w:pStyle | .//w:rStyle | .//w:tblStyle", namespaces=NS
    ):
        source_id = node.get(qn(W_NS, "val"))
        if source_id in style_id_map:
            node.set(qn(W_NS, "val"), style_id_map[source_id])

    referenced = {
        node.get(qn(W_NS, "val"))
        for node in cloned_block.xpath(".//w:pStyle | .//w:rStyle | .//w:tblStyle", namespaces=NS)
        if node.get(qn(W_NS, "val"))
    }
    existing = {
        node.get(qn(W_NS, "styleId"))
        for node in target_styles.xpath("./w:style", namespaces=NS)
    }
    pending = list(referenced)
    while pending:
        style_id = pending.pop()
        if style_id in existing:
            continue
        source_match = source_styles.xpath(
            "./w:style[@w:styleId=$style_id]",
            namespaces=NS,
            style_id=style_id,
        )
        if not source_match:
            continue
        style = deepcopy(source_match[0])
        for dependency in style.xpath("./w:basedOn | ./w:next | ./w:link", namespaces=NS):
            dependency_id = dependency.get(qn(W_NS, "val"))
            if dependency_id in style_id_map:
                dependency.set(qn(W_NS, "val"), style_id_map[dependency_id])
        # A merged Word package may contain only one default style for each
        # style type.  The retained report uses different IDs for its Normal,
        # Default Paragraph Font, and Normal Table styles; keep those styles
        # for dependency resolution but never import a second default flag.
        default_attribute = qn(W_NS, "default")
        style_type = style.get(qn(W_NS, "type"))
        if style.get(default_attribute) in {"1", "true", "on"}:
            target_default = target_styles.xpath(
                "./w:style[@w:type=$style_type and "
                "(@w:default='1' or @w:default='true' or @w:default='on')]",
                namespaces=NS,
                style_type=style_type,
            )
            if target_default:
                style.attrib.pop(default_attribute, None)
        target_styles.append(style)
        existing.add(style_id)
        for dependency in style.xpath("./w:basedOn | ./w:next | ./w:link", namespaces=NS):
            dependency_id = dependency.get(qn(W_NS, "val"))
            if dependency_id and dependency_id not in existing:
                pending.append(dependency_id)


def enforce_single_default_styles(styles_root: etree._Element) -> None:
    """Defensively guarantee one Word default style per style type."""
    default_attribute = qn(W_NS, "default")
    for style_type in ("paragraph", "character", "table", "numbering"):
        defaults = styles_root.xpath(
            "./w:style[@w:type=$style_type and "
            "(@w:default='1' or @w:default='true' or @w:default='on')]",
            namespaces=NS,
            style_type=style_type,
        )
        for duplicate in defaults[1:]:
            duplicate.attrib.pop(default_attribute, None)


def rewrite_missing_package_style_references(
    target_root: Path,
    reference_styles: etree._Element,
    target_styles: etree._Element,
) -> None:
    """Map source-only style IDs inside imported headers/footers to target IDs."""
    target_ids = {
        style.get(qn(W_NS, "styleId"))
        for style in target_styles.xpath("./w:style", namespaces=NS)
    }
    target_by_name_type = {
        (
            style.get(qn(W_NS, "type"), ""),
            style.xpath("string(w:name/@w:val)", namespaces=NS),
        ): style.get(qn(W_NS, "styleId"))
        for style in target_styles.xpath("./w:style", namespaces=NS)
    }
    source_to_target: dict[str, str] = {}
    for style in reference_styles.xpath("./w:style", namespaces=NS):
        source_id = style.get(qn(W_NS, "styleId"))
        key = (
            style.get(qn(W_NS, "type"), ""),
            style.xpath("string(w:name/@w:val)", namespaces=NS),
        )
        target_id = target_by_name_type.get(key)
        if source_id and target_id:
            source_to_target[source_id] = target_id

    for path in (target_root / "word").rglob("*.xml"):
        if path.name == "styles.xml":
            continue
        try:
            tree = etree.parse(str(path))
        except etree.XMLSyntaxError:
            continue
        changed = False
        for node in tree.getroot().xpath(
            ".//w:pStyle | .//w:rStyle | .//w:tblStyle", namespaces=NS
        ):
            style_id = node.get(qn(W_NS, "val"))
            if style_id not in target_ids and style_id in source_to_target:
                node.set(qn(W_NS, "val"), source_to_target[style_id])
                changed = True
        if changed:
            tree.write(
                str(path),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )


def normalize_word_property_order_in_package(target_root: Path) -> None:
    """Normalize edited Word property blocks so Word never has to repair them."""
    property_orders = {
        qn(W_NS, "rPr"): W_RPR_ORDER,
        qn(W_NS, "pPr"): W_PPR_ORDER,
        qn(W_NS, "tblPr"): W_TBLPR_ORDER,
        qn(W_NS, "tcPr"): W_TCPR_ORDER,
    }
    for path in (target_root / "word").rglob("*.xml"):
        try:
            tree = etree.parse(str(path))
        except etree.XMLSyntaxError:
            continue
        changed = False
        for element in tree.getroot().iter():
            order = property_orders.get(element.tag)
            if order is None or len(element) < 2:
                continue
            before = list(element)
            reorder_word_properties(element, order)
            changed = changed or before != list(element)
        if changed:
            tree.write(
                str(path),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )


def remove_orphan_bookmarks(document_root: etree._Element) -> int:
    """Remove bookmark endpoints orphaned by replacing template content ranges."""
    bookmark_id = qn(W_NS, "id")
    starts = document_root.xpath("//w:bookmarkStart", namespaces=NS)
    ends = document_root.xpath("//w:bookmarkEnd", namespaces=NS)
    start_ids = {node.get(bookmark_id) for node in starts}
    end_ids = {node.get(bookmark_id) for node in ends}
    removed = 0
    for node in starts:
        if node.get(bookmark_id) not in end_ids:
            node.getparent().remove(node)
            removed += 1
    for node in ends:
        if node.get(bookmark_id) not in start_ids:
            node.getparent().remove(node)
            removed += 1
    return removed


def load_reference_chart_captions(reference_path: Path) -> list[str]:
    """Return chart captions in chart order from the retained example report."""
    with ZipFile(reference_path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    captions: list[str] = []
    for chart in root.xpath("//c:chart", namespaces=NS):
        tables = chart.xpath("ancestor::w:tbl[1]", namespaces=NS)
        cells = chart.xpath("ancestor::w:tc[1]", namespaces=NS)
        caption_text = ""
        if tables and cells:
            table = tables[0]
            chart_cell = cells[0]
            chart_row = chart_cell.getparent()
            column_index = chart_row.findall("w:tc", namespaces=NS).index(chart_cell)
            for row in table.findall("w:tr", namespaces=NS):
                row_cells = row.findall("w:tc", namespaces=NS)
                if column_index >= len(row_cells):
                    continue
                candidates = row_cells[column_index].xpath(
                    ".//w:p[.//w:instrText[contains(., 'SEQ')]]",
                    namespaces=NS,
                )
                if candidates:
                    caption_text = element_text(candidates[0])
                    break
        title = re.sub(r"^[\s'▪]*图\s*\d*\s*[：:]\s*", "", caption_text).strip().lstrip("：: ")
        captions.append(f"图：{title}" if title else "")
    return captions


def map_reference_chart_captions(
    body_content: list[dict],
    reference_captions: list[str],
) -> dict[str, str]:
    """Align raw editable charts to the example, tolerating image-only charts."""
    charts = [
        item for item in body_content
        if item.get("type") == "figure" and item.get("object_type") == "chart"
    ]

    def normalized(value: str | None) -> str:
        value = re.sub(r"^[\s'▪]*图\s*\d*\s*[：:]*\s*", "", value or "")
        return re.sub(r"[\s，。、“”‘’：:；;（）()\-]", "", value)

    reference_titles = [normalized(value) for value in reference_captions]
    anchors: list[tuple[int, int]] = []
    reference_cursor = 0
    for raw_index, item in enumerate(charts):
        title = normalized(item.get("caption"))
        if not title:
            continue
        match_index = next(
            (
                index for index in range(reference_cursor, len(reference_titles))
                if title == reference_titles[index]
                or title in reference_titles[index]
                or reference_titles[index] in title
            ),
            None,
        )
        if match_index is not None:
            anchors.append((raw_index, match_index))
            reference_cursor = match_index + 1

    mapping: dict[str, str] = {}
    boundaries = [(-1, -1), *anchors, (len(charts), len(reference_captions))]
    for (raw_left, ref_left), (raw_right, ref_right) in zip(boundaries, boundaries[1:]):
        raw_gap = raw_right - raw_left - 1
        ref_gap = ref_right - ref_left - 1
        if raw_gap <= 0 or ref_gap < raw_gap:
            continue
        # Extra reference charts in a gap are image-only in the raw source and
        # occur before the next matching editable chart.
        offset = ref_gap - raw_gap
        for step in range(1, raw_gap + 1):
            raw_item = charts[raw_left + step]
            if raw_item.get("caption"):
                continue
            ref_index = ref_left + offset + step
            if 0 <= ref_index < len(reference_captions):
                mapping[raw_item.get("asset_id", "")] = reference_captions[ref_index]
    return mapping


def append_risk_section(
    target_body: etree._Element,
    insertion_before: etree._Element,
    risk_text: str,
    reference_document: etree._Element,
    reference_styles: etree._Element,
    target_styles: etree._Element,
    importer: PackagePartImporter | None,
    drawing_ids: itertools.count,
    risk_config: dict | None = None,
    risk_heading_text: str = "一、风险提示",
) -> None:
    """Append the report-ending risk section using the reference typography."""
    reference_body = reference_document.find("w:body", namespaces=NS)
    reference_breaks = body_section_breaks(reference_body)
    if len(reference_breaks) < 2:
        raise ValueError("Reference report has no risk/back-matter section break")
    break_paragraph = reference_breaks[1]
    children = list(reference_body)
    break_index = children.index(break_paragraph)
    heading_index = next(
        (
            index for index in range(break_index - 1, -1, -1)
            if re.match(r"^[一二三四五六七八九十]+、风险提示$", element_text(children[index]))
        ),
        None,
    )
    if heading_index is None:
        raise ValueError("Reference risk heading not found")
    item_prototypes = [
        child for child in children[heading_index + 1:break_index + 1]
        if child.tag == qn(W_NS, "p") and element_text(child)
    ]
    if not item_prototypes:
        raise ValueError("Reference risk item prototypes not found")

    content = re.sub(r"^风险提示\s*[：:]\s*", "", risk_text.strip())
    items = [value.strip(" ；;。") for value in re.split(r"[；;]", content) if value.strip(" ；;。")]
    if not items:
        items = ["请参阅报告正文所列风险因素"]

    blocks: list[etree._Element] = []
    heading = (
        clone_with_imported_relationships(children[heading_index], importer, drawing_ids)
        if importer is not None
        else deepcopy(children[heading_index])
    )
    set_paragraph_text(heading, risk_heading_text)
    heading_properties = ensure_paragraph_properties(heading)
    outline_level = heading_properties.find("w:outlineLvl", namespaces=NS)
    if outline_level is None:
        outline_level = etree.SubElement(heading_properties, qn(W_NS, "outlineLvl"))
    outline_level.set(qn(W_NS, "val"), "0")
    reorder_word_properties(heading_properties, W_PPR_ORDER)
    config = risk_config or {}
    for run in heading.findall("w:r", namespaces=NS):
        set_run_font(
            run,
            config.get("title_font", "思源黑体 CN Medium"),
            int(config.get("title_font_size_half_points", 22)),
        )
    blocks.append(heading)
    for index, value in enumerate(items):
        is_last = index == len(items) - 1
        prototype = item_prototypes[-1] if is_last else item_prototypes[min(index, len(item_prototypes) - 2)]
        paragraph = (
            clone_with_imported_relationships(prototype, importer, drawing_ids)
            if importer is not None
            else deepcopy(prototype)
        )
        set_paragraph_text(paragraph, f"{value}{'。' if is_last else '；'}")
        for run in paragraph.findall("w:r", namespaces=NS):
            set_run_font(
                run,
                config.get("item_font", "思源黑体 CN Medium"),
                int(config.get("item_font_size_half_points", 17)),
            )
        blocks.append(paragraph)
    for block in blocks:
        import_missing_styles(block, reference_styles, target_styles)
    insertion_index = target_body.index(insertion_before)
    for offset, block in enumerate(blocks):
        target_body.insert(insertion_index + offset, block)


def replace_back_matter(
    target_body: etree._Element,
    target_start: etree._Element,
    reference_document: etree._Element,
    importer: PackagePartImporter,
    drawing_ids: itertools.count,
    reference_styles: etree._Element,
    target_styles: etree._Element,
    section_break_ordinal: int,
) -> None:
    """Replace the closing report pages with the fixed reference section."""
    reference_body = reference_document.find("w:body", namespaces=NS)
    reference_breaks = body_section_breaks(reference_body)
    if section_break_ordinal < 1 or section_break_ordinal > len(reference_breaks):
        raise ValueError("Invalid fixed back-matter section break locator")
    reference_start = reference_breaks[section_break_ordinal - 1]
    target_children = list(target_body)
    for child in target_children[target_children.index(target_start):]:
        target_body.remove(child)
    reference_children = list(reference_body)
    # The generated final risk item already carries this section break; append
    # only the fixed analyst/legal pages that follow it.
    for child in reference_children[reference_children.index(reference_start) + 1:]:
        clone = clone_with_imported_relationships(child, importer, drawing_ids)
        import_missing_styles(clone, reference_styles, target_styles)
        target_body.append(clone)


def build_body_blocks(
    body_content: list[dict],
    prototypes: dict[str, etree._Element],
    assets: dict[str, dict],
    source_document: etree._Element,
    importer: PackagePartImporter,
    drawing_ids: itertools.count,
    source_styles: etree._Element,
    target_styles: etree._Element,
    layout: dict,
    chart_caption_fallbacks: dict[str, str] | None = None,
) -> list[etree._Element]:
    blocks: list[etree._Element] = []
    source_body = source_document.find("w:body", namespaces=NS)
    sequence_counts = {"图": 0, "表": 0}
    text_config = layout.get("text", {})
    chart_caption_fallbacks = chart_caption_fallbacks or {}

    def figure_caption(item: dict) -> str | None:
        explicit = item.get("caption")
        if explicit:
            return explicit
        if item.get("object_type") != "chart":
            return None
        return chart_caption_fallbacks.get(item.get("asset_id", ""))

    def make_caption(raw_text: str, explicit_kind: str | None = None) -> etree._Element:
        sequence_name = explicit_kind or ("表" if raw_text.lstrip().startswith("表") else "图")
        sequence_counts[sequence_name] += 1
        role = "table_caption" if sequence_name == "表" else "figure_caption"
        separator_config = deepcopy(layout.get("caption_separator", {}))
        if sequence_name == "表":
            separator_config["after_twips"] = separator_config.get(
                "table_after_twips", 120
            )
        return numbered_caption_paragraph(
            prototypes[role],
            raw_text,
            sequence_name,
            sequence_counts[sequence_name],
            separator_config,
        )

    def imported_figure(item: dict) -> etree._Element:
        asset = assets.get(item.get("asset_id"))
        if not asset:
            raise ValueError(f"Figure has no native asset: {item}")
        source_block = source_body[asset["body_index"]]
        visual = source_visual_paragraph(
            source_block, asset["type"], asset.get("object_index", 0)
        )
        cloned = clone_with_imported_relationships(visual, importer, drawing_ids)
        paragraph_properties = cloned.find("w:pPr", namespaces=NS)
        if paragraph_properties is not None:
            # Source drawing paragraphs may reference source-only styles such
            # as Normal (Web).  The fixed wrapper owns visual positioning, so
            # retaining that style creates an invalid dangling style ID.
            for child_name in ("w:pStyle", "w:numPr"):
                child = paragraph_properties.find(child_name, namespaces=NS)
                if child is not None:
                    paragraph_properties.remove(child)
        remove_picture_outline(cloned)
        return cloned

    def source_table(item: dict) -> etree._Element:
        asset = assets.get(item.get("asset_id"))
        if not asset:
            raise ValueError(f"Table has no native asset: {item.get('asset_id')}")
        table = source_body[asset["body_index"]]
        if table.tag != qn(W_NS, "tbl"):
            raise ValueError(f"Table locator does not point to w:tbl: {asset}")
        return table

    def imported_table(
        item: dict,
        nested: bool = True,
        table_config: dict | None = None,
    ) -> etree._Element:
        source_block = source_table(item)
        cloned = clone_with_imported_relationships(
            source_block, importer, drawing_ids
        )
        import_missing_styles(cloned, source_styles, target_styles)
        config = table_config or layout.get("figure_layout", {})
        normalize_native_table(
            cloned,
            int(config.get("cell_width_dxa", 8051)),
            0 if nested else int(config.get("indent_dxa", 2689)),
            text_config,
            layout.get("table_style") if nested else None,
        )
        return cloned

    def restyle_visual_container_table(
        table: etree._Element,
        groups: list[dict],
    ) -> etree._Element:
        """Apply report caption/source styling inside a figure-holder table."""
        rows = table.findall("w:tr", namespaces=NS)
        config = layout.get("figure_layout", {})
        table_width = int(config.get("width_dxa", 8051))
        table_indent = int(config.get("indent_dxa", 2689))
        max_width = int(config.get("max_drawing_width_emu", 5112000))

        properties = table.find("w:tblPr", namespaces=NS)
        if properties is None:
            properties = etree.Element(qn(W_NS, "tblPr"))
            table.insert(0, properties)

        # These figures arrive inside a one-column holder table.  Match the
        # retained report grid exactly: 14.2 cm wide with the template's
        # 4.74 cm table indent.  The visual itself is right-aligned separately
        # below when its native canvas is narrower than the holder.
        width = properties.find("w:tblW", namespaces=NS)
        if width is None:
            width = etree.SubElement(properties, qn(W_NS, "tblW"))
        width.set(qn(W_NS, "w"), str(table_width))
        width.set(qn(W_NS, "type"), "dxa")
        indent = properties.find("w:tblInd", namespaces=NS)
        if indent is None:
            indent = etree.SubElement(properties, qn(W_NS, "tblInd"))
        indent.set(qn(W_NS, "w"), str(table_indent))
        indent.set(qn(W_NS, "type"), "dxa")
        alignment = properties.find("w:jc", namespaces=NS)
        if alignment is not None:
            properties.remove(alignment)
        fixed = properties.find("w:tblLayout", namespaces=NS)
        if fixed is None:
            fixed = etree.SubElement(properties, qn(W_NS, "tblLayout"))
        fixed.set(qn(W_NS, "type"), "fixed")

        grid_columns = table.xpath("./w:tblGrid/w:gridCol", namespaces=NS)
        if len(grid_columns) == 1:
            grid_columns[0].set(qn(W_NS, "w"), str(table_width))
        for cell in table.xpath("./w:tr/w:tc", namespaces=NS):
            cell_properties = cell.find("w:tcPr", namespaces=NS)
            if cell_properties is None:
                cell_properties = etree.Element(qn(W_NS, "tcPr"))
                cell.insert(0, cell_properties)
            cell_width = cell_properties.find("w:tcW", namespaces=NS)
            if cell_width is None:
                cell_width = etree.SubElement(cell_properties, qn(W_NS, "tcW"))
            cell_width.set(qn(W_NS, "w"), str(table_width))
            cell_width.set(qn(W_NS, "type"), "dxa")
            reorder_word_properties(cell_properties, W_TCPR_ORDER)

        borders = properties.find("w:tblBorders", namespaces=NS)
        if borders is None:
            borders = etree.SubElement(properties, qn(W_NS, "tblBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = borders.find(f"w:{side}", namespaces=NS)
            if border is None:
                border = etree.SubElement(borders, qn(W_NS, side))
            border.set(qn(W_NS, "val"), "nil")
        reorder_word_properties(properties, W_TBLPR_ORDER)

        def first_cell(row: etree._Element) -> etree._Element:
            cell = row.find("w:tc", namespaces=NS)
            if cell is None:
                raise ValueError("Visual container row has no table cell")
            return cell

        def replace_cell_body(cell: etree._Element, block: etree._Element) -> None:
            for child in list(cell):
                if child.tag != qn(W_NS, "tcPr"):
                    cell.remove(child)
            cell.append(block)

        for group in groups:
            caption_row_index = int(group["caption_row"])
            visual_row_index = int(group["visual_row"])
            source_row_index = int(group["source_row"])
            if source_row_index >= len(rows):
                raise ValueError(f"Invalid visual table group: {group}")

            raw_caption = group.get("caption", "")
            explicit_kind = (
                "表" if raw_caption.lstrip(" '▪").startswith("表") else "图"
            )
            replace_cell_body(
                first_cell(rows[caption_row_index]),
                make_caption(raw_caption, explicit_kind),
            )

            visual_cell = first_cell(rows[visual_row_index])
            remove_picture_outline(visual_cell)
            for paragraph in visual_cell.xpath(".//w:p[.//w:drawing or .//w:pict]", namespaces=NS):
                set_paragraph_alignment(
                    paragraph, config.get("visual_alignment", "right")
                )
                scale_visual_to_width(paragraph, max_width)

            replace_cell_body(
                first_cell(rows[source_row_index]),
                data_source_paragraph(layout, group.get("source")),
            )

            for row_index in (
                caption_row_index, visual_row_index, source_row_index
            ):
                row_properties = rows[row_index].find("w:trPr", namespaces=NS)
                if row_properties is None:
                    row_properties = etree.Element(qn(W_NS, "trPr"))
                    rows[row_index].insert(0, row_properties)
                if row_properties.find("w:cantSplit", namespaces=NS) is None:
                    etree.SubElement(row_properties, qn(W_NS, "cantSplit"))

        return table

    index = 0
    while index < len(body_content):
        item = body_content[index]
        block_type = item.get("type")

        # The raw report uses this editorial instruction at the two places
        # where the finished reference uses a two-column visual component.
        if (
            block_type == "paragraph"
            and "两张图横向排列" in item.get("text", "")
        ):
            cursor = index + 1
            pending_caption: tuple[str, str] | None = None
            pair: list[tuple[etree._Element | None, etree._Element]] = []
            while cursor < len(body_content) and len(pair) < 2:
                candidate = body_content[cursor]
                candidate_type = candidate.get("type")
                if candidate_type in {"figure_caption", "table_caption"}:
                    kind = "表" if candidate_type == "table_caption" else "图"
                    pending_caption = (candidate.get("text", ""), kind)
                    cursor += 1
                    continue
                if candidate_type != "figure":
                    break
                raw_caption = figure_caption(candidate)
                explicit_kind = None
                if pending_caption is not None:
                    raw_caption, explicit_kind = pending_caption
                caption = (
                    make_caption(raw_caption, explicit_kind)
                    if raw_caption else None
                )
                pair.append((caption, imported_figure(candidate)))
                pending_caption = None
                cursor += 1
            if len(pair) == 2:
                blocks.append(
                    double_figure_layout_table(
                        [pair[0][0], pair[1][0]],
                        [pair[0][1], pair[1][1]],
                        layout,
                    )
                )
                index = cursor
                continue
            # An editorial layout instruction is never report prose.
            index += 1
            continue

        if block_type in {"heading1", "heading2", "heading3", "paragraph"}:
            paragraph = paragraph_from_prototype(
                prototypes[block_type],
                item.get("text", ""),
                item.get("runs") if block_type == "paragraph" else None,
                text_config if block_type == "paragraph" else None,
            )
            if block_type.startswith("heading"):
                level = int(block_type[-1]) - 1
                properties = ensure_paragraph_properties(paragraph)
                outline = properties.find("w:outlineLvl", namespaces=NS)
                if outline is None:
                    outline = etree.SubElement(properties, qn(W_NS, "outlineLvl"))
                outline.set(qn(W_NS, "val"), str(level))
                reorder_word_properties(properties, W_PPR_ORDER)
            blocks.append(paragraph)
            index += 1
            continue
        if block_type in {"figure_caption", "table_caption"}:
            raw_caption = item.get("text", "")
            explicit_kind = "表" if block_type == "table_caption" else "图"
            if (
                index + 1 < len(body_content)
                and body_content[index + 1].get("type") == "figure"
            ):
                visual = imported_figure(body_content[index + 1])
                blocks.append(
                    figure_layout_table(
                        make_caption(raw_caption, explicit_kind),
                        visual,
                        layout,
                    )
                )
                index += 2
                continue
            if (
                block_type == "table_caption"
                and index + 1 < len(body_content)
                and body_content[index + 1].get("type") == "table"
            ):
                table_item = body_content[index + 1]
                table_config = table_layout_config(source_table(table_item), layout)
                table = imported_table(table_item, table_config=table_config)
                blocks.append(
                    figure_layout_table(
                        make_caption(raw_caption, explicit_kind),
                        table,
                        layout,
                        table_config,
                    )
                )
                index += 2
                continue
            blocks.append(make_caption(raw_caption, explicit_kind))
            index += 1
            continue
        if block_type == "figure":
            caption = figure_caption(item)
            caption_node = (
                make_caption(caption)
                if caption else None
            )
            imported_visual = imported_figure(item)
            blocks.append(
                figure_layout_table(caption_node, imported_visual, layout)
            )
            index += 1
            continue
        if block_type == "table":
            visual_groups = item.get("visual_groups", [])
            if visual_groups:
                blocks.append(
                    restyle_visual_container_table(
                        imported_table(item, nested=False), visual_groups
                    )
                )
            else:
                table_config = table_layout_config(source_table(item), layout)
                blocks.append(
                    figure_layout_table(
                        None,
                        imported_table(item, table_config=table_config),
                        layout,
                        table_config,
                    )
                )
            index += 1
            continue
        # Unknown parser extensions are ignored only when they contain no text.
        if item.get("text"):
            blocks.append(
                paragraph_from_prototype(
                    prototypes["paragraph"],
                    item["text"],
                    item.get("runs"),
                    text_config,
                )
            )
        index += 1
    return blocks


def disable_update_fields_on_open(target_root: Path) -> None:
    """Avoid Word's misleading external-file warning on document open.

    macOS Word shows that warning for any package with ``updateFields=true``,
    even when the only fields are a TOC and page numbers.  The generator writes
    the TOC's cached entries itself; users can still choose Update Field in Word
    when they need freshly paginated page numbers.
    """
    path = target_root / "word/settings.xml"
    tree = etree.parse(str(path))
    root = tree.getroot()
    node = root.find("w:updateFields", namespaces=NS)
    if node is not None:
        root.remove(node)
        tree.write(
            str(path), xml_declaration=True, encoding="UTF-8", standalone=True
        )

    # The template's TOC prototypes mark their begin field characters dirty.
    # That flag alone is enough for Word for Mac to show the same warning even
    # without ``w:updateFields`` or any external relationship.  Cached TOC text
    # is already populated by the generator, so clear only the dirty flag.
    for part_path in (target_root / "word").rglob("*.xml"):
        part_tree = etree.parse(str(part_path))
        changed = False
        for field in part_tree.getroot().xpath("//w:fldChar", namespaces=NS):
            dirty = qn(W_NS, "dirty")
            if dirty in field.attrib:
                field.attrib.pop(dirty)
                changed = True
        if changed:
            part_tree.write(
                str(part_path),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )


def set_core_title(target_root: Path, title: str) -> None:
    path = target_root / "docProps/core.xml"
    tree = etree.parse(str(path))
    root = tree.getroot()
    node = root.find(qn(DC_NS, "title"))
    if node is None:
        node = etree.SubElement(root, qn(DC_NS, "title"))
    node.text = title
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


def fill_running_headers(target_root: Path, layout: dict) -> None:
    """Replace template control labels in retained running headers."""
    header_config = layout.get("running_header", {})
    category = header_config.get("category", "投资策略")
    report_type = header_config.get("report_type", "专题报告")
    for path in sorted((target_root / "word").glob("header*.xml")):
        tree = etree.parse(str(path))
        changed = False
        for paragraph in tree.getroot().xpath("//w:p", namespaces=NS):
            value = element_text(paragraph)
            compact = re.sub(r"\s+", "", value)
            if compact in {"[Table_PageText1]", "[Table_PageText2]"}:
                set_paragraph_text(paragraph, "")
                changed = True
            elif "XXXX|XX报告" in compact:
                set_paragraph_text(paragraph, f"{category} | {report_type}")
                changed = True
        if changed:
            tree.write(
                str(path),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )


def package_source_for_rels(rels_path: str) -> str:
    path = PurePosixPath(rels_path)
    if rels_path == "_rels/.rels":
        return ""
    parent = path.parent.parent
    return str(parent / path.name.removesuffix(".rels"))


def detach_external_chart_workbooks(root: Path) -> None:
    """Remove inaccessible Excel links while retaining native chart XML/cache."""
    relationship_attributes = {
        qn(R_NS, "id"), qn(R_NS, "embed"), qn(R_NS, "link")
    }
    for rels_path in root.rglob("*.rels"):
        rels_tree = etree.parse(str(rels_path))
        external_rids: set[str] = set()
        for rel in list(rels_tree.getroot()):
            rel_type = rel.get("Type", "")
            if (
                rel.get("TargetMode") == "External"
                and rel_type.endswith("/oleObject")
            ):
                if rel.get("Id"):
                    external_rids.add(rel.get("Id"))
                rels_tree.getroot().remove(rel)
        if not external_rids:
            continue

        package_name = rels_path.relative_to(root).as_posix()
        source_part = package_source_for_rels(package_name)
        part_path = root / source_part if source_part else None
        if part_path is not None and part_path.is_file() and part_path.suffix == ".xml":
            part_tree = etree.parse(str(part_path))
            for node in list(part_tree.getroot().iter()):
                referenced = {
                    node.get(attribute)
                    for attribute in relationship_attributes
                    if node.get(attribute)
                }
                if not referenced.intersection(external_rids):
                    continue
                if etree.QName(node).localname == "externalData":
                    parent = node.getparent()
                    if parent is not None:
                        parent.remove(node)
                else:
                    for attribute in relationship_attributes:
                        if node.get(attribute) in external_rids:
                            node.attrib.pop(attribute, None)
            part_tree.write(
                str(part_path),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
        rels_tree.write(
            str(rels_path),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )


def validate_package(root: Path) -> None:
    etree.parse(str(root / "word/document.xml"))
    missing: list[str] = []
    for rels_path in root.rglob("*.rels"):
        package_name = rels_path.relative_to(root).as_posix()
        source_part = package_source_for_rels(package_name)
        tree = etree.parse(str(rels_path))
        for rel in tree.getroot():
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target")
            resolved = (
                target.lstrip("/")
                if target.startswith("/")
                else posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))
            )
            if not (root / resolved).is_file():
                missing.append(f"{package_name}:{rel.get('Id')} -> {resolved}")
    if missing:
        raise ValueError("Broken internal relationships:\n" + "\n".join(missing))

    dangling_references: list[str] = []
    relationship_attributes = {
        qn(R_NS, "id"), qn(R_NS, "embed"), qn(R_NS, "link")
    }
    for part_path in root.rglob("*.xml"):
        if part_path.name == "[Content_Types].xml" or "_rels" in part_path.parts:
            continue
        package_name = part_path.relative_to(root).as_posix()
        tree = etree.parse(str(part_path))
        referenced_ids = {
            value
            for node in tree.getroot().iter()
            for attribute, value in node.attrib.items()
            if attribute in relationship_attributes and value
        }
        if not referenced_ids:
            continue
        rels_path = part_path.parent / "_rels" / f"{part_path.name}.rels"
        available_ids: set[str] = set()
        if rels_path.is_file():
            rels_tree = etree.parse(str(rels_path))
            available_ids = {
                rel.get("Id") for rel in rels_tree.getroot() if rel.get("Id")
            }
        for relationship_id in sorted(referenced_ids - available_ids):
            dangling_references.append(f"{package_name} -> {relationship_id}")
    if dangling_references:
        raise ValueError(
            "Dangling relationship references:\n" + "\n".join(dangling_references)
        )

    document_root = etree.parse(str(root / "word/document.xml")).getroot()
    bookmark_id = qn(W_NS, "id")
    bookmark_starts = {
        node.get(bookmark_id)
        for node in document_root.xpath("//w:bookmarkStart", namespaces=NS)
    }
    bookmark_ends = {
        node.get(bookmark_id)
        for node in document_root.xpath("//w:bookmarkEnd", namespaces=NS)
    }
    if bookmark_starts != bookmark_ends:
        raise ValueError(
            "Unbalanced bookmarks: "
            f"starts_only={sorted(bookmark_starts - bookmark_ends)}, "
            f"ends_only={sorted(bookmark_ends - bookmark_starts)}"
        )

    text = (root / "word/document.xml").read_text(encoding="utf-8")
    residues = [token for token in ("[Table_", "正文文案", "图表标题") if token in text]
    if residues:
        raise ValueError(f"Template placeholder residue: {residues}")


def zip_directory(root: Path, output_path: Path) -> None:
    temporary = output_path.with_suffix(".tmp.docx")
    if temporary.exists():
        temporary.unlink()
    with ZipFile(temporary, "w", ZIP_DEFLATED) as archive:
        for path in sorted(root.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(root).as_posix())
    os.replace(temporary, output_path)


def generate(
    structure_path: Path,
    template_path: Path,
    manifest_path: Path,
    layout_path: Path,
    output_dir: Path,
    source_path: Path | None = None,
) -> Path:
    structure = read_json(structure_path)
    manifest = read_json(manifest_path)
    layout = read_json(layout_path)
    source_path = source_path or Path(
        structure.get("source_docx") or manifest.get("source_docx") or ""
    )
    if not source_path.is_file():
        raise FileNotFoundError(f"Source DOCX not found: {source_path}")

    fallback_title = clean_filename_title(structure["filename"])
    title, summary, body_content = split_front_matter(
        structure.get("content", []), fallback_title
    )
    risk_heading = risk_heading_for_content(body_content)
    output_path = output_path_for(structure["filename"], output_dir)

    with tempfile.TemporaryDirectory(prefix="report-generator-") as temp:
        temp_path = Path(temp)
        source_root = temp_path / "source"
        target_root = temp_path / "target"
        source_root.mkdir()
        target_root.mkdir()
        with ZipFile(source_path) as archive:
            archive.extractall(source_root)
        with ZipFile(template_path) as archive:
            archive.extractall(target_root)

        target_document_tree = etree.parse(str(target_root / "word/document.xml"))
        target_document = target_document_tree.getroot()
        template_document = deepcopy(target_document)
        target_styles_tree = etree.parse(str(target_root / "word/styles.xml"))
        target_styles = target_styles_tree.getroot()
        source_document = etree.parse(str(source_root / "word/document.xml")).getroot()
        source_styles = etree.parse(str(source_root / "word/styles.xml")).getroot()

        prototypes = find_prototypes(target_document, target_styles, layout["styles"])
        fill_cover(target_document, target_styles, layout, title, summary)
        body = target_document.find("w:body", namespaces=NS)
        name_to_id, _ = style_maps(target_styles)
        body_start_style_id = name_to_id[layout["body"]["start_style"]]
        body_start = next(
            child for child in body
            if child.tag == qn(W_NS, "p") and paragraph_style_id(child) == body_start_style_id
        )
        # A missing source caption stays missing.  Never borrow a title from
        # a sample report or from instructional charts in the template.
        chart_caption_fallbacks: dict[str, str] = {}
        replace_index(
            body,
            body_start,
            layout,
            body_content,
            risk_heading,
            chart_caption_fallbacks,
        )

        section_breaks = body_section_breaks(body)
        end_ordinal = layout["body"]["end_before_section_break"]
        body_end = section_breaks[end_ordinal - 1]
        children = list(body)
        start_index = children.index(body_start)
        end_index = children.index(body_end)
        for child in children[start_index:end_index]:
            body.remove(child)

        importer = PackagePartImporter(source_root, target_root)
        assets = {item["asset_id"]: item for item in manifest.get("assets", [])}
        existing_ids = [
            int(value)
            for value in target_document.xpath("//wp:docPr/@id", namespaces={"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"})
            if str(value).isdigit()
        ]
        drawing_ids = itertools.count(max(existing_ids, default=0) + 1)
        new_blocks = build_body_blocks(
            body_content,
            prototypes,
            assets,
            source_document,
            importer,
            drawing_ids,
            source_styles,
            target_styles,
            layout,
            chart_caption_fallbacks,
        )
        insertion_index = body.index(body_end)
        for offset, block in enumerate(new_blocks):
            body.insert(insertion_index + offset, block)

        # The analyst and legal pages already live after ``body_end`` in the
        # retained template and remain untouched in place.
        importer.save()
        risk_text = next(
            (
                item.get("text", "") for item in summary
                if item.get("text", "").lstrip().startswith("风险提示：")
            ),
            layout.get("risk_section", {}).get("fallback_text", ""),
        )
        append_risk_section(
            body,
            body_end,
            risk_text,
            template_document,
            target_styles,
            target_styles,
            None,
            drawing_ids,
            layout.get("risk_section"),
            risk_heading,
        )
        clear_template_markers(target_document)

        enforce_single_default_styles(target_styles)
        remove_orphan_bookmarks(target_document)
        target_document_tree.write(
            str(target_root / "word/document.xml"),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        target_styles_tree.write(
            str(target_root / "word/styles.xml"),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        chart_template_value = layout.get("chart_template")
        chart_template_path = (
            ROOT_DIR / chart_template_value
            if chart_template_value else DEFAULT_CHART_TEMPLATE
        )
        apply_chart_template(target_root, chart_template_path)
        normalize_word_property_order_in_package(target_root)
        disable_update_fields_on_open(target_root)
        set_core_title(target_root, title)
        fill_running_headers(target_root, layout)
        detach_external_chart_workbooks(target_root)
        validate_package(target_root)
        zip_directory(target_root, output_path)

    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--structure", type=Path, default=DEFAULT_STRUCTURE)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--assets", type=Path, default=DEFAULT_MANIFEST)
    parser.add_argument("--layout", type=Path, default=DEFAULT_LAYOUT)
    parser.add_argument("--source", type=Path)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output = generate(
        args.structure,
        args.template,
        args.assets,
        args.layout,
        args.output_dir,
        args.source,
    )
    print(f"Generated: {output}")


if __name__ == "__main__":
    main()
