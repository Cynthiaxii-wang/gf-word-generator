"""OOXML package helpers for copying editable Word objects.

The functions in this module work below the ``python-docx`` abstraction.  A
chart is not only a ``<c:chart>`` node: it is a relationship-backed graph of
chart XML, chart styles, colour styles, drawings, media and (sometimes)
embedded workbooks.  ``PackagePartImporter`` copies that graph recursively and
rewrites relationship targets when package part names collide.
"""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path, PurePosixPath
import mimetypes
import posixpath
import re
from typing import Dict, Iterable, Optional

from lxml import etree


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"


def relationship_part_name(part_name: str) -> str:
    """Return the relationship part path for an OPC part."""

    path = PurePosixPath(part_name)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def resolve_target(source_part: str, target: str) -> str:
    """Resolve an internal relationship target to a package-absolute path."""

    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relative_target(source_part: str, target_part: str) -> str:
    """Return *target_part* relative to the directory of *source_part*."""

    return posixpath.relpath(target_part, posixpath.dirname(source_part))


class PackagePartImporter:
    """Copy relationship-backed parts from one extracted DOCX into another."""

    def __init__(self, source_root: Path, target_root: Path):
        self.source_root = Path(source_root)
        self.target_root = Path(target_root)
        self.source_content_types = etree.parse(
            str(self.source_root / "[Content_Types].xml")
        )
        self.target_content_types = etree.parse(
            str(self.target_root / "[Content_Types].xml")
        )
        self._part_map: Dict[str, str] = {}
        self._rid_map: Dict[str, str] = {}
        self._counter = 1

        self.document_rels_path = self.target_root / "word/_rels/document.xml.rels"
        self.document_rels = etree.parse(str(self.document_rels_path))
        self.source_document_rels = etree.parse(
            str(self.source_root / "word/_rels/document.xml.rels")
        )

    @staticmethod
    def _relationship_by_id(tree: etree._ElementTree, rid: str) -> etree._Element:
        matches = tree.getroot().xpath(
            "./pr:Relationship[@Id=$rid]",
            namespaces={"pr": REL_NS},
            rid=rid,
        )
        if not matches:
            raise KeyError(f"Relationship not found: {rid}")
        return matches[0]

    def _next_document_rid(self) -> str:
        used = {
            rel.get("Id")
            for rel in self.document_rels.getroot()
            if rel.get("Id")
        }
        numeric = [
            int(match.group(1))
            for rid in used
            if (match := re.fullmatch(r"rId(\d+)", rid))
        ]
        candidate = max(numeric, default=0) + 1
        while f"rId{candidate}" in used:
            candidate += 1
        return f"rId{candidate}"

    def import_document_relationship(self, source_rid: str) -> str:
        """Import one relationship from source ``document.xml.rels``."""

        if source_rid in self._rid_map:
            return self._rid_map[source_rid]

        source_rel = self._relationship_by_id(self.source_document_rels, source_rid)
        new_rid = self._next_document_rid()
        new_rel = etree.Element(f"{{{REL_NS}}}Relationship")
        new_rel.set("Id", new_rid)
        new_rel.set("Type", source_rel.get("Type"))

        if source_rel.get("TargetMode") == "External":
            new_rel.set("Target", source_rel.get("Target"))
            new_rel.set("TargetMode", "External")
        else:
            source_part = resolve_target("word/document.xml", source_rel.get("Target"))
            target_part = self.import_part(source_part)
            new_rel.set("Target", relative_target("word/document.xml", target_part))

        self.document_rels.getroot().append(new_rel)
        self._rid_map[source_rid] = new_rid
        return new_rid

    def _allocate_target_part(self, source_part: str) -> str:
        source_path = PurePosixPath(source_part)
        preferred = source_part
        if not (self.target_root / preferred).exists():
            return preferred

        while True:
            stem = source_path.stem
            suffix = source_path.suffix
            name = f"rg_{stem}_{self._counter}{suffix}"
            self._counter += 1
            candidate = str(source_path.parent / name)
            if not (self.target_root / candidate).exists():
                return candidate

    def import_part(self, source_part: str) -> str:
        """Recursively copy an internal OPC part and its relationships."""

        source_part = source_part.lstrip("/")
        if source_part in self._part_map:
            return self._part_map[source_part]

        source_path = self.source_root / source_part
        if not source_path.is_file():
            raise FileNotFoundError(f"Missing source package part: {source_part}")

        target_part = self._allocate_target_part(source_part)
        self._part_map[source_part] = target_part
        target_path = self.target_root / target_part
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(source_path.read_bytes())
        self._copy_content_type(source_part, target_part)

        source_rels_name = relationship_part_name(source_part)
        source_rels_path = self.source_root / source_rels_name
        if source_rels_path.is_file():
            rel_tree = etree.parse(str(source_rels_path))
            external_rids: set[str] = set()
            for rel in list(rel_tree.getroot()):
                if rel.get("TargetMode") == "External":
                    # Publicly generated reports must be self-contained.  Excel
                    # charts copied from research drafts often retain a link to
                    # the analyst's local/OneDrive workbook.  Copying that link
                    # makes Word show an unreadable-content/external-field
                    # warning on every other computer.  Keep the native chart
                    # and its cached series, but detach the inaccessible source.
                    if rel.get("Id"):
                        external_rids.add(rel.get("Id"))
                    rel_tree.getroot().remove(rel)
                    continue
                dependency = resolve_target(source_part, rel.get("Target"))
                imported_dependency = self.import_part(dependency)
                rel.set("Target", relative_target(target_part, imported_dependency))

            if external_rids and target_path.suffix.lower() == ".xml":
                part_tree = etree.parse(str(target_path))
                relationship_attributes = {
                    f"{{{R_NS}}}id",
                    f"{{{R_NS}}}embed",
                    f"{{{R_NS}}}link",
                }
                for node in list(part_tree.getroot().iter()):
                    referenced = {
                        node.get(attribute)
                        for attribute in relationship_attributes
                        if node.get(attribute)
                    }
                    if not referenced.intersection(external_rids):
                        continue
                    # c:externalData is only a pointer to the linked workbook;
                    # removing it does not flatten or rasterise the chart.
                    if etree.QName(node).localname == "externalData":
                        parent = node.getparent()
                        if parent is not None:
                            parent.remove(node)
                    else:
                        for attribute in relationship_attributes:
                            if node.get(attribute) in external_rids:
                                node.attrib.pop(attribute, None)
                part_tree.write(
                    str(target_path),
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

            target_rels_path = self.target_root / relationship_part_name(target_part)
            target_rels_path.parent.mkdir(parents=True, exist_ok=True)
            rel_tree.write(
                str(target_rels_path),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
        return target_part

    def _content_type_for(self, part_name: str) -> tuple[str, bool]:
        root = self.source_content_types.getroot()
        override = root.xpath(
            "./ct:Override[@PartName=$name]",
            namespaces={"ct": CT_NS},
            name=f"/{part_name}",
        )
        if override:
            return override[0].get("ContentType"), True

        extension = PurePosixPath(part_name).suffix.lstrip(".")
        default = root.xpath(
            "./ct:Default[@Extension=$ext]",
            namespaces={"ct": CT_NS},
            ext=extension,
        )
        if default:
            return default[0].get("ContentType"), False

        guessed = mimetypes.guess_type(part_name)[0] or "application/octet-stream"
        return guessed, False

    def _copy_content_type(self, source_part: str, target_part: str) -> None:
        content_type, source_was_override = self._content_type_for(source_part)
        root = self.target_content_types.getroot()
        extension = PurePosixPath(target_part).suffix.lstrip(".")
        defaults = root.xpath(
            "./ct:Default[@Extension=$ext]",
            namespaces={"ct": CT_NS},
            ext=extension,
        )
        if defaults and defaults[0].get("ContentType") == content_type and not source_was_override:
            return

        existing = root.xpath(
            "./ct:Override[@PartName=$name]",
            namespaces={"ct": CT_NS},
            name=f"/{target_part}",
        )
        if existing:
            return
        node = etree.Element(f"{{{CT_NS}}}Override")
        node.set("PartName", f"/{target_part}")
        node.set("ContentType", content_type)
        root.append(node)

    def save(self) -> None:
        self.document_rels.write(
            str(self.document_rels_path),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        self.target_content_types.write(
            str(self.target_root / "[Content_Types].xml"),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )


def clone_with_imported_relationships(
    element: etree._Element,
    importer: PackagePartImporter,
    next_drawing_id: Optional[Iterable[int]] = None,
) -> etree._Element:
    """Deep-copy an OOXML block and import every document relationship it uses."""

    clone = deepcopy(element)
    relationship_attributes = {
        f"{{{R_NS}}}id",
        f"{{{R_NS}}}embed",
        f"{{{R_NS}}}link",
    }
    for node in clone.iter():
        for attribute in relationship_attributes:
            source_rid = node.get(attribute)
            if source_rid:
                node.set(attribute, importer.import_document_relationship(source_rid))

    if next_drawing_id is not None:
        ids = iter(next_drawing_id)
        for node in clone.xpath(".//wp:docPr | .//pic:cNvPr", namespaces={"wp": WP_NS, "pic": PIC_NS}):
            node.set("id", str(next(ids)))
    return clone
