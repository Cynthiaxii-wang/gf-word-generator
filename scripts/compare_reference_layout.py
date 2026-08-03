"""Print compact OOXML evidence for TOC and visual-layout comparison."""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"w": W, "c": C, "a": A}


def attr(local: str) -> str:
    return f"{{{W}}}{local}"


def text(node: etree._Element) -> str:
    return "".join(node.xpath(".//w:t/text()", namespaces=NS)).strip()


def field_text(node: etree._Element) -> str:
    return " | ".join(
        value.strip()
        for value in node.xpath(".//w:instrText/text()", namespaces=NS)
        if value.strip()
    )


def table_geometry(table: etree._Element) -> str:
    width = table.xpath("string(w:tblPr/w:tblW/@w:w)", namespaces=NS)
    indent = table.xpath("string(w:tblPr/w:tblInd/@w:w)", namespaces=NS)
    grid = table.xpath("w:tblGrid/w:gridCol/@w:w", namespaces=NS)
    rows = len(table.findall("w:tr", namespaces=NS))
    return f"w={width or '-'} ind={indent or '-'} grid={','.join(grid) or '-'} rows={rows}"


def paragraph_style(paragraph: etree._Element) -> str:
    node = paragraph.find("w:pPr/w:pStyle", namespaces=NS)
    return node.get(attr("val"), "-") if node is not None else "-"


def inspect(path: Path) -> None:
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", namespaces=NS)
    print(f"\n=== {path.name} ===")
    print("INDEX/FIELD BLOCKS")
    for index, child in enumerate(body):
        values = field_text(child)
        visible = text(child)
        if values or visible in {"目录索引", "图表索引"}:
            kind = etree.QName(child).localname
            style = paragraph_style(child) if kind == "p" else "-"
            print(f"{index:03d} {kind} style={style} field={values!r} text={visible[:180]!r}")

    print("VISUAL CONTAINERS")
    ordinal = 0
    for index, child in enumerate(body):
        drawings = child.xpath(".//w:drawing", namespaces=NS)
        if not drawings:
            continue
        ordinal += len(drawings)
        charts = len(child.xpath(".//c:chart", namespaces=NS))
        images = len(child.xpath(".//a:blip", namespaces=NS))
        kind = etree.QName(child).localname
        geometry = table_geometry(child) if kind == "tbl" else "standalone"
        visible = " ".join(text(child).split())
        print(
            f"body={index:03d} visual_through={ordinal:02d} {kind} {geometry} "
            f"draw={len(drawings)} chart={charts} image={images} text={visible[:220]!r}"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", nargs="+", type=Path)
    args = parser.parse_args()
    for path in args.docx:
        inspect(path)


if __name__ == "__main__":
    main()
