"""Summarize native chart objects in an XLSX/DOCX OOXML package."""

from __future__ import annotations

import argparse
import posixpath
from zipfile import ZipFile

from lxml import etree


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def resolve(base: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def rels_name(part: str) -> str:
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", f"{name}.rels")


def relationships(zf: ZipFile, part: str) -> dict[str, str]:
    name = rels_name(part)
    if name not in zf.namelist():
        return {}
    root = etree.fromstring(zf.read(name))
    return {
        rel.get("Id"): resolve(part, rel.get("Target"))
        for rel in root.xpath("./pr:Relationship", namespaces=NS)
        if rel.get("TargetMode") != "External"
    }


def chart_summary(zf: ZipFile, part: str) -> str:
    root = etree.fromstring(zf.read(part))
    plot = root.find(".//c:plotArea", namespaces=NS)
    chart_types = []
    if plot is not None:
        for child in plot:
            if child.tag.startswith(f"{{{NS['c']}}}") and child.tag.endswith("Chart"):
                chart_types.append(etree.QName(child).localname)
    title = "".join(root.xpath(".//c:title//a:t/text()", namespaces=NS)).strip()
    legend = root.find(".//c:legend/c:legendPos", namespaces=NS)
    legend_pos = legend.get("val") if legend is not None else "none"
    fonts = sorted(set(root.xpath(".//a:defRPr/@typeface | .//a:rPr/@typeface", namespaces=NS)))
    sizes = sorted(set(root.xpath(".//a:defRPr/@sz | .//a:rPr/@sz", namespaces=NS)))
    colors = sorted(set(root.xpath(".//a:srgbClr/@val", namespaces=NS)))
    widths = sorted(set(root.xpath(".//a:ln/@w", namespaces=NS)))
    return (
        f"{part}: types={','.join(chart_types)} title={title!r} legend={legend_pos} "
        f"fonts={fonts} sizes={sizes} colors={colors} line_widths={widths}"
    )


def inspect_xlsx(path: str) -> None:
    with ZipFile(path) as zf:
        workbook_part = "xl/workbook.xml"
        workbook = etree.fromstring(zf.read(workbook_part))
        workbook_rels = relationships(zf, workbook_part)
        for sheet in workbook.xpath(".//m:sheets/m:sheet", namespaces=NS):
            sheet_name = sheet.get("name")
            sheet_part = workbook_rels.get(sheet.get(f"{{{NS['r']}}}id"))
            if not sheet_part:
                continue
            sheet_root = etree.fromstring(zf.read(sheet_part))
            sheet_rels = relationships(zf, sheet_part)
            drawing_nodes = sheet_root.xpath(".//m:drawing", namespaces=NS)
            print(f"\n[{sheet_name}] part={sheet_part}")
            for drawing_node in drawing_nodes:
                drawing_part = sheet_rels.get(drawing_node.get(f"{{{NS['r']}}}id"))
                if not drawing_part:
                    continue
                drawing = etree.fromstring(zf.read(drawing_part))
                drawing_rels = relationships(zf, drawing_part)
                for anchor in drawing.xpath("./xdr:twoCellAnchor | ./xdr:oneCellAnchor", namespaces=NS):
                    chart = anchor.find(".//c:chart", namespaces=NS)
                    if chart is None:
                        continue
                    chart_part = drawing_rels.get(chart.get(f"{{{NS['r']}}}id"))
                    from_col = anchor.findtext("xdr:from/xdr:col", namespaces=NS)
                    from_row = anchor.findtext("xdr:from/xdr:row", namespaces=NS)
                    to_col = anchor.findtext("xdr:to/xdr:col", namespaces=NS)
                    to_row = anchor.findtext("xdr:to/xdr:row", namespaces=NS)
                    print(f"  anchor=({from_col},{from_row})->({to_col},{to_row})")
                    if chart_part:
                        print("  " + chart_summary(zf, chart_part))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("xlsx")
    args = parser.parse_args()
    inspect_xlsx(args.xlsx)
