# PPTX QuickTool

PPTX QuickTool is a company-internal desktop tool for generating PowerPoint training document skeletons.

The first product slice focuses on creating the project foundation. Later tickets will add the document model, PPTX page planning, generation, navigation links, and the full form workflow.

## Requirements

- Python 3.11 or newer
- Tkinter support in the local Python installation
- `python-pptx` for PowerPoint generation

## Run The Desktop App

```powershell
python scripts/run_app.py
```

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
