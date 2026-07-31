from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx_quicktool import __version__


DIST = ROOT / "dist"
BUILD = ROOT / "build"
PYINSTALLER_DIST = BUILD / "pyinstaller-dist"
ARCHIVE_DIR = DIST / "archive"
RELEASE_NAME = f"PPTX_QuickTool_v{__version__}"
RELEASE_DIR = DIST / RELEASE_NAME
RELEASE_ARCHIVE = DIST / f"{RELEASE_NAME}.zip"
EXE_NAME = "PPTX QuickTool"


def archive_stale_release_artifacts() -> None:
    DIST.mkdir(exist_ok=True)
    ARCHIVE_DIR.mkdir(exist_ok=True)

    for directory in DIST.glob("PPTX_QuickTool_v*"):
        if directory.is_dir() and directory.name != RELEASE_NAME:
            archive_path = ARCHIVE_DIR / f"{directory.name}.zip"
            _zip_directory(directory, archive_path, root=DIST)
            shutil.rmtree(directory)

    for archive in DIST.glob("PPTX_QuickTool_v*.zip"):
        if archive.name != RELEASE_ARCHIVE.name:
            archived_zip = ARCHIVE_DIR / archive.name
            if archived_zip.exists():
                archived_zip.unlink()
            shutil.move(str(archive), archived_zip)

    stale_pyinstaller_dist = DIST / EXE_NAME
    if stale_pyinstaller_dist.exists():
        shutil.rmtree(stale_pyinstaller_dist)


def build_executable() -> Path:
    try:
        import PyInstaller.__main__
    except ModuleNotFoundError as error:
        raise SystemExit(
            "PyInstaller is required to build the Windows executable. "
            "Install build dependencies with: python -m pip install -e .[build]"
        ) from error

    data_separator = ";" if sys.platform.startswith("win") else ":"
    template_source = ROOT / "src" / "pptx_quicktool" / "assets" / "default-training-document-template.pptx"
    template_target = "pptx_quicktool/assets"
    if PYINSTALLER_DIST.exists():
        shutil.rmtree(PYINSTALLER_DIST)

    PyInstaller.__main__.run(
        [
            "--noconfirm",
            "--clean",
            "--windowed",
            "--name",
            EXE_NAME,
            "--distpath",
            str(PYINSTALLER_DIST),
            "--workpath",
            str(BUILD),
            "--paths",
            str(ROOT / "src"),
            "--add-data",
            f"{template_source}{data_separator}{template_target}",
            str(ROOT / "scripts" / "run_app.py"),
        ]
    )

    executable = PYINSTALLER_DIST / EXE_NAME / f"{EXE_NAME}.exe"
    if not executable.exists():
        raise SystemExit(f"Expected executable was not created: {executable}")
    return executable


def build_release_package() -> Path:
    archive_stale_release_artifacts()
    executable = build_executable()

    if RELEASE_DIR.exists():
        shutil.rmtree(RELEASE_DIR)
    RELEASE_DIR.mkdir(parents=True)

    app_dir = RELEASE_DIR / EXE_NAME
    shutil.copytree(executable.parent, app_dir)
    shutil.copy2(ROOT / "README.md", RELEASE_DIR / "README.md")
    docs_dir = RELEASE_DIR / "docs"
    docs_dir.mkdir()
    shutil.copy2(ROOT / "docs" / "user-guide.md", docs_dir / "user-guide.md")

    if RELEASE_ARCHIVE.exists():
        RELEASE_ARCHIVE.unlink()
    _zip_directory(RELEASE_DIR, RELEASE_ARCHIVE, root=RELEASE_DIR.parent)
    return RELEASE_ARCHIVE


def _zip_directory(directory: Path, archive_path: Path, root: Path) -> None:
    archive_path.parent.mkdir(parents=True, exist_ok=True)
    if archive_path.exists():
        archive_path.unlink()
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in directory.rglob("*"):
            if path.is_file():
                archive.write(path, path.relative_to(root).as_posix())


if __name__ == "__main__":
    output = build_release_package()
    print(output)
