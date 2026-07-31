from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
BUILD = ROOT / "build"
RELEASE_NAME = "PPTX_QuickTool_v0.1.0"
RELEASE_DIR = DIST / RELEASE_NAME
RELEASE_ARCHIVE = DIST / f"{RELEASE_NAME}.zip"
EXE_NAME = "PPTX QuickTool"


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

    PyInstaller.__main__.run(
        [
            "--noconfirm",
            "--clean",
            "--windowed",
            "--name",
            EXE_NAME,
            "--distpath",
            str(DIST),
            "--workpath",
            str(BUILD),
            "--paths",
            str(ROOT / "src"),
            "--add-data",
            f"{template_source}{data_separator}{template_target}",
            str(ROOT / "scripts" / "run_app.py"),
        ]
    )

    executable = DIST / EXE_NAME / f"{EXE_NAME}.exe"
    if not executable.exists():
        raise SystemExit(f"Expected executable was not created: {executable}")
    return executable


def build_release_package() -> Path:
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
    with zipfile.ZipFile(RELEASE_ARCHIVE, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in RELEASE_DIR.rglob("*"):
            if path.is_file():
                archive.write(path, path.relative_to(RELEASE_DIR.parent).as_posix())
    return RELEASE_ARCHIVE


if __name__ == "__main__":
    output = build_release_package()
    print(output)
