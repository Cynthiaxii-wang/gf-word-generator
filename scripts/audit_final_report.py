"""Compact structural audit for the generated 广发 report."""

from __future__ import annotations

import argparse
import posixpath
from zipfile import ZipFile

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def text(node):
    return "".join(node.xpath(".//w:t/text()", namespaces=NS))


def audit(path: str) -> None:
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
        captions = root.xpath(
            ".//w:p[.//w:instrText[contains(., 'SEQ 图') or contains(., 'SEQ 表')]]",
            namespaces=NS,
        )
        print("captions", len(captions))
        print("caption_prefixes", sorted({text(p)[:2] for p in captions}))
        print(
            "caption_sizes",
            sorted(set(root.xpath(
                ".//w:p[.//w:instrText[contains(., 'SEQ 图') or contains(., 'SEQ 表')]]"
                "//w:rPr/w:sz/@w:val",
                namespaces=NS,
            ))),
        )
        print(
            "caption_fonts",
            sorted(set(root.xpath(
                ".//w:p[.//w:instrText[contains(., 'SEQ 图') or contains(., 'SEQ 表')]]"
                "//w:rPr/w:rFonts/@w:eastAsia",
                namespaces=NS,
            ))),
        )
        print(
            "caption_keep_markers",
            len(root.xpath(
                ".//w:p[.//w:instrText[contains(., 'SEQ 图') or contains(., 'SEQ 表')]]"
                "/w:pPr/w:keepNext | "
                ".//w:p[.//w:instrText[contains(., 'SEQ 图') or contains(., 'SEQ 表')]]"
                "/w:pPr/w:keepLines",
                namespaces=NS,
            )),
        )
        print(
            "caption_line_rules",
            sorted(set(root.xpath(
                ".//w:p[.//w:instrText[contains(., 'SEQ 图') or contains(., 'SEQ 表')]]"
                "/w:pPr/w:spacing/@w:line",
                namespaces=NS,
            ))),
        )
        print(
            "light_runs",
            len(root.xpath(".//w:rFonts[@w:eastAsia='思源黑体 CN Light']", namespaces=NS)),
            "medium_runs",
            len(root.xpath(".//w:rFonts[@w:eastAsia='思源黑体 CN Medium']", namespaces=NS)),
        )
        wrappers = root.xpath(
            ".//w:tbl[w:tblPr/w:tblW[@w:w='8051'] and w:tblPr/w:tblInd[@w:w='2689']]",
            namespaces=NS,
        )
        nested_tables = root.xpath(".//w:tbl//w:tbl", namespaces=NS)
        print("position_wrappers", len(wrappers), "nested_native_tables", len(nested_tables))

        rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        rel_map = {
            rel.get("Id"): rel.get("Target")
            for rel in rels.xpath("./pr:Relationship", namespaces=NS)
        }
        chart_parts = []
        for node in root.xpath(".//c:chart", namespaces=NS):
            target = rel_map.get(node.get(f"{{{NS['r']}}}id"))
            if target:
                chart_parts.append(posixpath.normpath(posixpath.join("word", target)))
        chart_parts = list(dict.fromkeys(chart_parts))
        sizes = set()
        colors = set()
        latin_typefaces = set()
        east_asian_typefaces = set()
        for part in chart_parts:
            chart = etree.fromstring(archive.read(part))
            sizes.update(chart.xpath(".//a:defRPr/@sz | .//a:rPr/@sz", namespaces=NS))
            colors.update(chart.xpath(".//c:ser/c:spPr//a:srgbClr/@val", namespaces=NS))
            latin_typefaces.update(chart.xpath(".//a:latin/@typeface", namespaces=NS))
            east_asian_typefaces.update(chart.xpath(".//a:ea/@typeface", namespaces=NS))
        print("editable_charts", len(chart_parts))
        print("chart_text_sizes", sorted(sizes))
        print("chart_latin_typefaces", sorted(latin_typefaces))
        print("chart_east_asian_typefaces", sorted(east_asian_typefaces))
        print("chart_series_colors", sorted(colors))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    args = parser.parse_args()
    audit(args.docx)
