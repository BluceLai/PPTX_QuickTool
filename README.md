# PPTX QuickTool

PPTX QuickTool is a company-internal desktop tool for generating PowerPoint training document skeletons.

The tool uses the built-in training document PPTX template by default. Users can also select another PPTX template file before generation. Generated files include a cover slide, table of contents, section slides, optional content slides, table-of-contents links, and Chinese return-to-contents links on section pages. The app starts with a default `前言` section and wraps long agenda slides into up to three columns.

## Requirements

- Python 3.11 or newer
- Tkinter support in the local Python installation
- `python-pptx` for PowerPoint generation

## Run The Desktop App

```powershell
python scripts/run_app.py
```

Chinese usage instructions are available in [`docs/user-guide.md`](docs/user-guide.md).

For a non-interactive smoke run:

```powershell
python scripts/run_app.py --smoke
```

## Run Tests

```powershell
python -m unittest discover -s tests
```

## Build A Package Archive

```powershell
python scripts/build.py
```

The build command writes a source archive under `dist/`.

## Build A Windows EXE Release

Install the build dependency once:

```powershell
python -m pip install -e .[build]
```

Then build the distributable zip:

```powershell
python scripts/build_exe.py
```

The build command writes `dist/PPTX_QuickTool_v0.1.0.zip`. Share that zip with coworkers. They can extract it and run `PPTX QuickTool/PPTX QuickTool.exe`.
