"""asset_manager.py v3 - inventory editable objects in a DOCX package.

Unlike v2, the manifest is ordered by document occurrence, includes source
body/object locators, records native tables, and lists the complete recursive
dependency graph required to copy each chart or image without rasterising it.
"""

from __future__ import annotations

import argparse
import json
import posixpath
from collections import defaultdict
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

from lxml import etree


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_DIR = ROOT_DIR / "input"
DEFAULT_OUTPUT = ROOT_DIR / "runtime" / "assets_manifest.json"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "a": A_NS, "c": C_NS, "r": R_NS, "pr": PR_NS}


def rels_name(part_name: str) -> str:
    path = PurePosixPath(part_name)
    return str(path.parent / "_rels" / f"{path.name}.rels")


def resolve_target(part_name: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part_name), target))


def load_relationships(archive: ZipFile, part_name: str) -> dict[str, dict]:
    name = rels_name(part_name)
    if name not in archive.namelist():
        return {}
    root = etree.fromstring(archive.read(name))
    return {
        rel.get("Id"): {
            "target": rel.get("Target"),
            "type": rel.get("Type"),
            "target_mode": rel.get("TargetMode"),
        }
        for rel in root
    }


def dependency_closure(
    archive: ZipFile,
    source_part: str,
    seen: set[str] | None = None,
) -> list[str]:
    """Return all internal package parts needed by *source_part*."""

    seen = seen or set()
    source_part = source_part.lstrip("/")
    if source_part in seen:
        return []
    seen.add(source_part)
    result = [source_part]
    relationship_part = rels_name(source_part)
    if relationship_part in archive.namelist():
        result.append(relationship_part)
    for rel in load_relationships(archive, source_part).values():
        if rel["target_mode"] == "External":
            continue
        target = resolve_target(source_part, rel["target"])
        if target in archive.namelist():
            result.extend(dependency_closure(archive, target, seen))
    return result


def scan_docx(input_path: Path) -> dict:
    counters: defaultdict[str, int] = defaultdict(int)
    assets: list[dict] = []

    with ZipFile(input_path) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        relationships = load_relationships(archive, "word/document.xml")
        body = document.find(f"{{{W_NS}}}body")

        for body_index, block in enumerate(body):
            tag = etree.QName(block).localname
            if tag == "tbl":
                counters["table"] += 1
                assets.append(
                    {
                        "asset_id": f"table_{counters['table']:03d}",
                        "type": "table",
                        "source": "word/document.xml",
                        "body_index": body_index,
                        "object_index": 0,
                        "dependencies": [],
                    }
                )
                continue
            if tag != "p":
                continue

            visual_refs: list[tuple[str, str]] = []
            for chart in block.xpath(".//c:chart", namespaces=NS):
                rid = chart.get(f"{{{R_NS}}}id")
                if rid:
                    visual_refs.append(("chart", rid))
            for blip in block.xpath(".//a:blip", namespaces=NS):
                rid = blip.get(f"{{{R_NS}}}embed") or blip.get(f"{{{R_NS}}}link")
                if rid:
                    visual_refs.append(("image", rid))

            local_indexes: defaultdict[str, int] = defaultdict(int)
            for object_type, rid in visual_refs:
                object_index = local_indexes[object_type]
                local_indexes[object_type] += 1
                rel = relationships.get(rid)
                if not rel:
                    continue
                target_part = (
                    rel["target"]
                    if rel["target_mode"] == "External"
                    else resolve_target("word/document.xml", rel["target"])
                )
                counters[object_type] += 1
                dependencies = []
                if rel["target_mode"] != "External" and target_part in archive.namelist():
                    dependencies = dependency_closure(archive, target_part)
                assets.append(
                    {
                        "asset_id": f"{object_type}_{counters[object_type]:03d}",
                        "type": object_type,
                        "rid": rid,
                        "relationship_type": rel["type"],
                        "source": target_part,
                        "body_index": body_index,
                        "object_index": object_index,
                        "dependencies": dependencies,
                    }
                )

        # Keep embedded objects that are directly attached to document.xml even
        # if their host markup is unsupported by the current parser.
        referenced = {item.get("rid") for item in assets}
        for rid, rel in relationships.items():
            if rid in referenced or not any(
                token in rel["type"] for token in ("oleObject", "package")
            ):
                continue
            counters["embedding"] += 1
            target_part = resolve_target("word/document.xml", rel["target"])
            assets.append(
                {
                    "asset_id": f"embedding_{counters['embedding']:03d}",
                    "type": "embedding",
                    "rid": rid,
                    "relationship_type": rel["type"],
                    "source": target_part,
                    "body_index": None,
                    "object_index": None,
                    "dependencies": dependency_closure(archive, target_part),
                }
            )

    return {
        "version": 3,
        "filename": input_path.name,
        "source_docx": str(input_path.resolve()),
        "assets": assets,
    }


def parse_args() -> argparse.Namespace:
    candidates = sorted(DEFAULT_INPUT_DIR.glob("*.docx"))
    default_input = candidates[0] if candidates else None
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", nargs="?", type=Path, default=default_input)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.input is None or not args.input.is_file():
        raise SystemExit("No input DOCX found")
    result = scan_docx(args.input)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    counts = defaultdict(int)
    for asset in result["assets"]:
        counts[asset["type"]] += 1
    print(f"Scanned: {args.input}")
    print(f"Manifest: {args.output}")
    print("Assets:", dict(counts))


if __name__ == "__main__":
    main()
