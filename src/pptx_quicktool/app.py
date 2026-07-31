from __future__ import annotations

import tkinter as tk
from tkinter import ttk


APP_TITLE = "PPTX QuickTool"
APP_GEOMETRY = "960x640"


def create_main_window() -> tk.Tk:
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry(APP_GEOMETRY)
    root.minsize(760, 520)

    _configure_styles(root)
    _build_shell(root)

    return root


def _configure_styles(root: tk.Tk) -> None:
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")

    style.configure("App.TFrame", background="#f5f6f8")
    style.configure("Header.TLabel", background="#f5f6f8", font=("Segoe UI", 18, "bold"))
    style.configure("Body.TLabel", background="#f5f6f8", font=("Segoe UI", 10))
    style.configure("Status.TLabel", background="#eef1f5", font=("Segoe UI", 9))


def _build_shell(root: tk.Tk) -> None:
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    frame = ttk.Frame(root, padding=24, style="App.TFrame")
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(2, weight=1)

    title = ttk.Label(frame, text=APP_TITLE, style="Header.TLabel")
    title.grid(row=0, column=0, sticky="w")

    subtitle = ttk.Label(
        frame,
        text="Training document skeleton generator",
        style="Body.TLabel",
    )
    subtitle.grid(row=1, column=0, sticky="w", pady=(6, 20))

    preview = tk.Text(frame, height=12, wrap="word", borderwidth=1, relief="solid")
    preview.insert(
        "1.0",
        "Ticket 01 skeleton is ready.\n\n"
        "Upcoming workflow:\n"
        "1. Define the training document input model.\n"
        "2. Generate a deterministic page plan.\n"
        "3. Export a new PPTX skeleton.\n"
        "4. Add table-of-contents navigation links.\n",
    )
    preview.configure(state="disabled")
    preview.grid(row=2, column=0, sticky="nsew")

    status = ttk.Label(
        frame,
        text="Ready",
        anchor="w",
        padding=(8, 4),
        style="Status.TLabel",
    )
    status.grid(row=3, column=0, sticky="ew", pady=(16, 0))
