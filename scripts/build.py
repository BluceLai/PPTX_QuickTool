from __future__ import annotations

import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
ARCHIVE = DIST / "pptx-quicktool-source.zip"

INCLUDED_ROOT_FILES = [
    ".gitignore",
    "AGENTS.md",
    "README.md",
    "pyproject.toml",
]

INCLUDED_DIRS = [
    "docs",
    "scripts",
    "src",
    "tests",
]


def iter_package_files():
    for name in INCLUDED_ROOT_FILES:
        path = ROOT / name
        if path.exists():
            yield path

    for dirname in INCLUDED_DIRS:
        directory = ROOT / dirname
        if not directory.exists():
            continue
        for path in directory.rglob("*"):
            if path.is_file() and "__pycache__" not in path.parts:
                yield path


def build_archive() -> Path:
    DIST.mkdir(exist_ok=True)
    with zipfile.ZipFile(ARCHIVE, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in iter_package_files():
            archive.write(path, path.relative_to(ROOT).as_posix())
    return ARCHIVE


if __name__ == "__main__":
    output = build_archive()
    print(output)
