import argparse
import subprocess
import sys
from pathlib import Path

from utils import (
    BASE_DIR,
    UI_PKG_DIR,
    compile_resources,
    generate_qml_module_artifacts,
)


def build_parser() -> argparse.ArgumentParser:
    """Create the command-line argument parser."""
    parser = argparse.ArgumentParser(description="Build Power Narrator targets.")
    parser.add_argument(
        "target",
        nargs="?",
        choices=("ui", "pptx"),
        default="ui",
        help="Build target to compile (default: ui)",
    )
    return parser


def _build_args(target: str) -> list[str | Path]:
    """Return the Nuitka command arguments for the selected target."""
    args: list[str | Path] = [
        "uv",
        "run",
        "python",
        "-m",
        "nuitka",
        f"--output-dir={BASE_DIR / 'dist'}",
        "--include-data-files=src/power_narrator/pptx/resources/narration-icon.png=power_narrator/pptx/resources/narration-icon.png",
        f"--output-filename=power-narrator-{target}",
    ]

    if target == "ui":
        return [
            *args,
            "--enable-plugin=pyside6",
            "--include-qt-plugins=qml,multimedia",
            "--mode=app",
            UI_PKG_DIR,
        ]

    return [
        *args,
        "--mode=onefile",
        BASE_DIR / "src" / "power_narrator" / "cli",
    ]


def run_build(target: str = "ui"):
    """Build the selected target using Nuitka."""
    if target == "ui":
        generate_qml_module_artifacts()
        compile_resources()

    try:
        _ = subprocess.run(_build_args(target), check=True)
    except subprocess.CalledProcessError as e:
        print(f"Nuitka build failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    args = build_parser().parse_args()
    run_build(args.target)
