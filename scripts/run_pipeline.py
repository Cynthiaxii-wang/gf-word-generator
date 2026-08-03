"""
Run parse -> asset inventory -> clean -> template generation for one DOCX.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import subprocess
import sys
import uuid


ROOT_DIR = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = ROOT_DIR / "scripts"


def default_input() -> Path | None:
    candidates = sorted(
        (ROOT_DIR / "input").glob("*.docx")
    )
    return candidates[0] if candidates else None



def run(command: list[str]) -> None:

    print(
        "Running:",
        " ".join(map(str, command))
    )

    subprocess.run(
        command,
        cwd=ROOT_DIR,
        check=True
    )



def main():

    parser = argparse.ArgumentParser(
        description=__doc__
    )


    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        default=default_input()
    )


    parser.add_argument(
        "--work-dir",
        type=Path,
        default=None
    )


    parser.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT_DIR / "output"
    )


    parser.add_argument(
        "--template",
        type=Path,
        default=ROOT_DIR / "template" / "总量通用.docx"
    )


    args = parser.parse_args()



    if args.input is None or not args.input.is_file():

        raise SystemExit(
            f"No input DOCX found: {args.input}"
        )



    # ==========================
    # 独立运行目录
    # ==========================

    if args.work_dir is None:

        run_id = uuid.uuid4().hex[:8]

        args.work_dir = (
            ROOT_DIR
            /
            "runtime"
            /
            run_id
        )


    args.work_dir.mkdir(
        parents=True,
        exist_ok=True
    )


    args.output_dir.mkdir(
        parents=True,
        exist_ok=True
    )


    print(
        "Input:",
        args.input
    )

    print(
        "Work:",
        args.work_dir
    )

    print(
        "Output:",
        args.output_dir
    )



    raw = (
        args.work_dir
        /
        "raw_structure.json"
    )


    assets = (
        args.work_dir
        /
        "assets_manifest.json"
    )


    clean = (
        args.work_dir
        /
        "clean_structure.json"
    )


    python = sys.executable



    # ==========================
    # parse
    # ==========================


    run(
        [
            python,
            str(SCRIPTS_DIR / "parse_report.py"),
            str(args.input),
            "--output",
            str(raw)
        ]
    )



    # ==========================
    # asset
    # ==========================


    run(
        [
            python,
            str(SCRIPTS_DIR / "asset_manager.py"),
            str(args.input),
            "--output",
            str(assets)
        ]
    )



    # ==========================
    # clean
    # ==========================


    run(
        [
            python,
            str(SCRIPTS_DIR / "clean_structure.py"),

            "--raw",
            str(raw),

            "--assets",
            str(assets),

            "--output",
            str(clean),
        ]
    )



    # ==========================
    # generate
    # ==========================


    run(
        [
            python,

            str(SCRIPTS_DIR / "generate_report.py"),

            "--structure",
            str(clean),

            "--assets",
            str(assets),

            "--source",
            str(args.input),

            "--template",
            str(args.template),

            "--output-dir",
            str(args.output_dir),

        ]
    )


    print(
        "Pipeline finished."
    )



if __name__ == "__main__":

    main()