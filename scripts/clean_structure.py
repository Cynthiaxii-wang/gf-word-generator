"""Normalize parsed report blocks and attach native object identifiers."""

from __future__ import annotations

import argparse
import json
from collections import defaultdict, deque
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT_DIR / "examples" / "output"
DEFAULT_RAW = OUTPUT_DIR / "raw_structure.json"
DEFAULT_ASSETS = OUTPUT_DIR / "assets_manifest.json"
DEFAULT_OUTPUT = OUTPUT_DIR / "clean_structure.json"


def clean_text(text: str | None) -> str | None:
    if not text:
        return text
    text = text.strip()
    if len(text) % 2 == 0:
        half = len(text) // 2
        if text[:half] == text[half:]:
            return text[:half]
    return text


def build_asset_queues(manifest: dict) -> tuple[dict, deque]:
    visual_queues: dict[tuple[str, str], deque] = defaultdict(deque)
    table_queue: deque = deque()
    for asset in manifest.get("assets", []):
        if asset["type"] == "table":
            table_queue.append(asset)
        elif asset["type"] in {"chart", "image"}:
            key = (asset["type"], Path(asset["source"]).name)
            visual_queues[key].append(asset)
    return visual_queues, table_queue


def attach_assets(content: list[dict], manifest: dict) -> list[dict]:
    visual_queues, table_queue = build_asset_queues(manifest)
    result = []
    for item in content:
        item = dict(item)
        if "text" in item:
            original_text = item["text"]
            item["text"] = clean_text(original_text)
            if item["text"] != original_text:
                # A de-duplicated parser artefact no longer has a one-to-one
                # run map; fall back to an ordinary Light-weight paragraph.
                item.pop("runs", None)
        if item.get("type") == "figure":
            source_name = Path(item.get("path", "")).name
            key = (item.get("object_type"), source_name)
            queue = visual_queues.get(key)
            if queue:
                asset = queue.popleft()
                item["asset_id"] = asset["asset_id"]
            item.pop("path", None)
        elif item.get("type") == "table" and table_queue:
            item["asset_id"] = table_queue.popleft()["asset_id"]
        result.append(item)
    return result


def merge_captions(content: list[dict]) -> list[dict]:
    """Merge a caption into the following figure, preserving all figures."""

    result: list[dict] = []
    pending_caption: str | None = None
    for item in content:
        item_type = item.get("type")
        if item_type == "figure_caption":
            if pending_caption:
                result.append({"type": "figure_caption", "text": pending_caption})
            pending_caption = item.get("text", "")
            continue
        if item_type == "figure" and pending_caption:
            item = dict(item)
            item["caption"] = pending_caption
            pending_caption = None
            result.append(item)
            continue
        if pending_caption:
            # A parser artefact may expose text living inside the same drawing.
            # Only skip it when it is an exact doubled-string residue.
            text = item.get("text", "")
            if item_type == "paragraph" and text and len(text) % 2 == 0 and text[: len(text)//2] == text[len(text)//2 :]:
                continue
            result.append({"type": "figure_caption", "text": pending_caption})
            pending_caption = None
        result.append(item)
    if pending_caption:
        result.append({"type": "figure_caption", "text": pending_caption})
    return result


def clean(raw: dict, manifest: dict) -> dict:
    content = attach_assets(raw.get("content", []), manifest)
    content = merge_captions(content)
    return {
        "filename": raw["filename"],
        "source_docx": manifest.get("source_docx"),
        "content": content,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--raw", type=Path, default=DEFAULT_RAW)
    parser.add_argument("--assets", type=Path, default=DEFAULT_ASSETS)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    raw = json.loads(args.raw.read_text(encoding="utf-8"))
    manifest = json.loads(args.assets.read_text(encoding="utf-8"))
    output = clean(raw, manifest)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    missing = [
        item for item in output["content"]
        if item.get("type") in {"figure", "table"} and not item.get("asset_id")
    ]
    print(f"Wrote: {args.output}")
    print(f"Unmatched native objects: {len(missing)}")


if __name__ == "__main__":
    main()
