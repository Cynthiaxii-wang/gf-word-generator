"""Apply the 广发 Excel chart design language to editable Word charts."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
import posixpath
from zipfile import ZipFile

from lxml import etree


C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"c": C_NS, "a": A_NS, "r": R_NS, "pr": PR_NS}

BAR_AREA_PIE_PALETTE = [
    "B8BFE4", "474DB1", "2E3160", "808080", "FFCB8D", "F7B941",
    "EF9900", "AED9FF", "77B6E4", "397099", "8A8CC7", "D9D9D9",
]
LINE_PALETTE = [
    "7C80C8", "F7B941", "D9D9D9", "EF9900", "FFCB8D", "B8BFE4",
    "2E3160", "AED9FF", "77B6E4", "397099", "808080", "5B5FA2",
]

A_SPPR_ORDER = {
    name: index for index, name in enumerate((
        "xfrm", "prstGeom", "custGeom", "noFill", "solidFill", "gradFill",
        "blipFill", "pattFill", "grpFill", "ln", "effectLst", "effectDag",
        "scene3d", "sp3d", "extLst",
    ))
}
A_TEXT_RPR_ORDER = {
    name: index for index, name in enumerate((
        "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill",
        "grpFill", "effectLst", "effectDag", "highlight", "uLnTx", "uLn",
        "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick",
        "hlinkMouseOver", "rtl", "extLst",
    ))
}


def reorder_drawing_properties(parent: etree._Element, order: dict[str, int]) -> None:
    children = list(parent)
    parent[:] = sorted(
        children,
        key=lambda child: order.get(etree.QName(child).localname, len(order)),
    )


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


@dataclass
class TemplateGroup:
    kind: str
    bar_direction: str
    series_count: int
    root: etree._Element
    group: etree._Element


def chart_groups(root: etree._Element) -> list[etree._Element]:
    plot = root.find(".//c:plotArea", namespaces=NS)
    if plot is None:
        return []
    return [
        child
        for child in plot
        if child.tag.startswith(f"{{{C_NS}}}")
        and etree.QName(child).localname.endswith("Chart")
    ]


def bar_direction(group: etree._Element) -> str:
    node = group.find("c:barDir", namespaces=NS)
    return node.get("val", "") if node is not None else ""


def load_template_groups(path: Path) -> list[TemplateGroup]:
    groups: list[TemplateGroup] = []
    with ZipFile(path) as archive:
        names = sorted(
            name for name in archive.namelist()
            if name.startswith("xl/charts/chart") and name.endswith(".xml")
        )
        for name in names:
            root = etree.fromstring(archive.read(name))
            for group in chart_groups(root):
                groups.append(
                    TemplateGroup(
                        kind=etree.QName(group).localname,
                        bar_direction=bar_direction(group),
                        series_count=len(group.findall("c:ser", namespaces=NS)),
                        root=root,
                        group=group,
                    )
                )
    if not groups:
        raise ValueError(f"Chart template contains no native charts: {path}")
    return groups


def best_template_group(
    target: etree._Element,
    templates: list[TemplateGroup],
) -> TemplateGroup:
    target_kind = etree.QName(target).localname
    lookup_kind = "lineChart" if target_kind == "scatterChart" else target_kind
    target_series = len(target.findall("c:ser", namespaces=NS))
    target_bar_direction = bar_direction(target)
    candidates = [item for item in templates if item.kind == lookup_kind]
    if not candidates:
        candidates = templates
    return min(
        candidates,
        key=lambda item: (
            item.bar_direction != target_bar_direction
            if target_kind == "barChart" else False,
            abs(item.series_count - target_series),
            item.series_count < target_series,
        ),
    )


def replace_child(
    parent: etree._Element,
    tag: str,
    source: etree._Element | None,
) -> None:
    old = parent.find(tag, namespaces=NS)
    if old is not None:
        parent.remove(old)
    if source is not None:
        parent.append(deepcopy(source))


def apply_series_style(target_group: etree._Element, reference: TemplateGroup) -> None:
    target_series = target_group.findall("c:ser", namespaces=NS)
    reference_series = reference.group.findall("c:ser", namespaces=NS)
    if not reference_series:
        return
    for index, series in enumerate(target_series):
        source = reference_series[index % len(reference_series)]
        for child_name in ("c:spPr", "c:marker", "c:explosion", "c:smooth"):
            replace_child(series, child_name, source.find(child_name, namespaces=NS))

    # Geometry-only properties are safe to copy; grouping/bar direction are
    # data semantics and remain source-owned.
    for child_name in ("c:gapWidth", "c:overlap", "c:firstSliceAng", "c:holeSize"):
        source = reference.group.find(child_name, namespaces=NS)
        if source is not None:
            replace_child(target_group, child_name, source)


def copy_chart_level_style(target: etree._Element, reference: etree._Element) -> None:
    target_chart_space_sppr = target.find("c:spPr", namespaces=NS)
    reference_chart_space_sppr = reference.find("c:spPr", namespaces=NS)
    if target_chart_space_sppr is not None:
        target.remove(target_chart_space_sppr)
    if reference_chart_space_sppr is not None:
        target.append(deepcopy(reference_chart_space_sppr))

    target_chart = target.find("c:chart", namespaces=NS)
    reference_chart = reference.find("c:chart", namespaces=NS)
    if target_chart is None or reference_chart is None:
        return
    title = target_chart.find("c:title", namespaces=NS)
    if title is not None:
        target_chart.remove(title)
    deleted = target_chart.find("c:autoTitleDeleted", namespaces=NS)
    if deleted is None:
        deleted = etree.Element(qn(C_NS, "autoTitleDeleted"))
        target_chart.insert(0, deleted)
    deleted.set("val", "1")

    for child_name in ("c:legend", "c:plotVisOnly", "c:dispBlanksAs", "c:showDLblsOverMax"):
        replace_child(
            target_chart,
            child_name,
            reference_chart.find(child_name, namespaces=NS),
        )

    target_plot = target_chart.find("c:plotArea", namespaces=NS)
    reference_plot = reference_chart.find("c:plotArea", namespaces=NS)
    if target_plot is None or reference_plot is None:
        return
    replace_child(target_plot, "c:spPr", reference_plot.find("c:spPr", namespaces=NS))

    reference_axes: dict[str, list[etree._Element]] = {}
    for axis in reference_plot:
        local = etree.QName(axis).localname
        if local in {"catAx", "dateAx", "valAx", "serAx"}:
            reference_axes.setdefault(local, []).append(axis)
    axis_offsets: dict[str, int] = {}
    for axis in target_plot:
        local = etree.QName(axis).localname
        if local not in {"catAx", "dateAx", "valAx", "serAx"}:
            continue
        candidates = reference_axes.get(local) or reference_axes.get("valAx") or []
        if not candidates:
            continue
        offset = axis_offsets.get(local, 0)
        source_axis = candidates[min(offset, len(candidates) - 1)]
        axis_offsets[local] = offset + 1
        for child_name in (
            "c:majorGridlines", "c:minorGridlines", "c:spPr", "c:txPr",
            "c:majorTickMark", "c:minorTickMark", "c:tickLblPos",
        ):
            replace_child(axis, child_name, source_axis.find(child_name, namespaces=NS))


def replace_drawing_fill(parent: etree._Element, color: str) -> None:
    fills = {
        qn(A_NS, name)
        for name in ("noFill", "solidFill", "gradFill", "pattFill", "blipFill", "grpFill")
    }
    insertion = next(
        (index for index, child in enumerate(parent) if child.tag in fills),
        0,
    )
    for child in list(parent):
        if child.tag in fills:
            parent.remove(child)
    solid = etree.Element(qn(A_NS, "solidFill"))
    etree.SubElement(solid, qn(A_NS, "srgbClr")).set("val", color)
    parent.insert(insertion, solid)


def ensure_series_properties(series: etree._Element) -> etree._Element:
    properties = series.find("c:spPr", namespaces=NS)
    if properties is not None:
        return properties
    properties = etree.Element(qn(C_NS, "spPr"))
    children = list(series)
    insert_after = max(
        (
            index for index, child in enumerate(children)
            if etree.QName(child).localname in {"idx", "order", "tx"}
        ),
        default=-1,
    )
    series.insert(insert_after + 1, properties)
    return properties


def apply_spec_palette(root: etree._Element) -> None:
    for group in chart_groups(root):
        kind = etree.QName(group).localname
        palette = LINE_PALETTE if kind in {"lineChart", "scatterChart"} else BAR_AREA_PIE_PALETTE
        for index, series in enumerate(group.findall("c:ser", namespaces=NS)):
            color = palette[index % len(palette)]
            properties = ensure_series_properties(series)
            if kind in {"lineChart", "scatterChart"}:
                line = properties.find("a:ln", namespaces=NS)
                if line is None:
                    line = etree.SubElement(properties, qn(A_NS, "ln"))
                replace_drawing_fill(line, color)
            else:
                replace_drawing_fill(properties, color)
                line = properties.find("a:ln", namespaces=NS)
                if line is not None:
                    replace_drawing_fill(line, color)
            reorder_drawing_properties(properties, A_SPPR_ORDER)


def set_chart_typography(
    root: etree._Element,
    chinese_font: str = "思源黑体 CN Light",
    latin_font: str = "Arial",
    size_hundredth_points: int = 600,
    color: str = "000000",
) -> None:
    for props in root.xpath(".//a:defRPr | .//a:rPr | .//a:endParaRPr", namespaces=NS):
        props.set("sz", str(size_hundredth_points))
        props.set("b", "0")
        props.attrib.pop("typeface", None)
        for local, font in (("latin", latin_font), ("ea", chinese_font), ("cs", latin_font)):
            node = props.find(f"a:{local}", namespaces=NS)
            if node is None:
                node = etree.SubElement(props, qn(A_NS, local))
            node.set("typeface", font)
        replace_drawing_fill(props, color)
        reorder_drawing_properties(props, A_TEXT_RPR_ORDER)


def remove_chart_borders(root: etree._Element) -> None:
    """Remove editable chart and plot-area outlines without changing geometry."""
    properties_nodes = []
    chart_space_properties = root.find("c:spPr", namespaces=NS)
    if chart_space_properties is not None:
        properties_nodes.append(chart_space_properties)
    plot_properties = root.find(".//c:plotArea/c:spPr", namespaces=NS)
    if plot_properties is not None:
        properties_nodes.append(plot_properties)
    for properties in properties_nodes:
        line = properties.find("a:ln", namespaces=NS)
        if line is None:
            line = etree.SubElement(properties, qn(A_NS, "ln"))
        for child in list(line):
            if etree.QName(child).localname in {
                "noFill", "solidFill", "gradFill", "pattFill", "blipFill",
            }:
                line.remove(child)
        line.insert(0, etree.Element(qn(A_NS, "noFill")))
        reorder_drawing_properties(properties, A_SPPR_ORDER)


def apply_chart_style(
    chart_path: Path,
) -> None:
    tree = etree.parse(str(chart_path))
    root = tree.getroot()
    if not chart_groups(root):
        return
    # Sheet1 is the normative page: 6 pt Source Han Sans Light for Chinese,
    # 6 pt Arial for Latin/numbers, black text, and the published color order.
    # The later workbook sheets are examples only and are not cloned.
    apply_spec_palette(root)
    set_chart_typography(root)
    remove_chart_borders(root)
    tree.write(
        str(chart_path),
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def document_chart_parts(target_root: Path) -> list[str]:
    document = etree.parse(str(target_root / "word/document.xml")).getroot()
    rels = etree.parse(str(target_root / "word/_rels/document.xml.rels")).getroot()
    rel_map = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.xpath("./pr:Relationship", namespaces=NS)
        if rel.get("TargetMode") != "External"
    }
    parts = []
    for node in document.xpath(".//c:chart", namespaces=NS):
        target = rel_map.get(node.get(qn(R_NS, "id")))
        if target:
            parts.append(posixpath.normpath(posixpath.join("word", target)))
    return list(dict.fromkeys(parts))


def apply_chart_template(target_root: Path, template_path: Path) -> int:
    """Apply only the workbook's Sheet1 rules; images remain untouched."""

    if not template_path.is_file():
        raise FileNotFoundError(f"Chart template not found: {template_path}")
    parts = document_chart_parts(target_root)
    for part in parts:
        apply_chart_style(target_root / part)
    return len(parts)
