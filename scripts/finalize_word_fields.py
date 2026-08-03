"""Reassert deterministic typography after Microsoft Word updates fields."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def set_font(run: etree._Element, name: str, size: int | None = None) -> None:
    props = run.find("w:rPr", namespaces=NS)
    if props is None:
        props = etree.Element(qn("rPr"))
        run.insert(0, props)
    fonts = props.find("w:rFonts", namespaces=NS)
    if fonts is None:
        fonts = etree.Element(qn("rFonts"))
        props.insert(0, fonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(attr), name)
    if size is not None:
        for local in ("sz", "szCs"):
            node = props.find(f"w:{local}", namespaces=NS)
            if node is None:
                node = etree.SubElement(props, qn(local))
            node.set(qn("val"), str(size))
    for local in ("b", "bCs"):
        node = props.find(f"w:{local}", namespaces=NS)
        if node is None:
            node = etree.SubElement(props, qn(local))
        node.set(qn("val"), "0")


def finalize(input_path: Path, output_path: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="report-finalize-", dir="/private/tmp") as temp:
        root = Path(temp)
        with ZipFile(input_path) as archive:
            archive.extractall(root)
        document_path = root / "word/document.xml"
        tree = etree.parse(str(document_path))
        document = tree.getroot()

        for run in document.xpath(".//w:r[w:rPr/w:rFonts]", namespaces=NS):
            fonts = run.find("w:rPr/w:rFonts", namespaces=NS)
            names = {
                fonts.get(qn(attr), "")
                for attr in ("ascii", "hAnsi", "eastAsia", "cs")
            }
            if "思源黑体 CN Light" in names:
                set_font(run, "思源黑体 CN Light")
            elif "思源黑体 CN Medium" in names:
                set_font(run, "思源黑体 CN Medium")

        captions = document.xpath(
            ".//w:p[.//w:instrText[contains(., 'SEQ 图') or contains(., 'SEQ 表')]]",
            namespaces=NS,
        )
        for paragraph in captions:
            for run in paragraph.findall("w:r", namespaces=NS):
                set_font(run, "思源黑体 CN Medium", 16)

        tree.write(
            str(document_path),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        temporary = output_path.with_suffix(".tmp.docx")
        with ZipFile(temporary, "w", ZIP_DEFLATED) as archive:
            for path in sorted(root.rglob("*")):
                if path.is_file():
                    archive.write(path, path.relative_to(root).as_posix())
        os.replace(temporary, output_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    finalize(args.input, args.out or args.input)
