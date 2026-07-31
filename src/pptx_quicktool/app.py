from __future__ import annotations

import tkinter as tk
import re
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, ttk

from .page_plan import generate_page_plan
from .pptx_exporter import DEFAULT_TEMPLATE_PATH, export_page_plan_to_pptx
from .pptx_verifier import verify_pptx_output
from .training_document import (
    SectionInput,
    TrainingDocumentInput,
    validate_training_document_input,
)


APP_TITLE = "PPTX QuickTool"
APP_GEOMETRY = "1080x720"


@dataclass
class EditableSection:
    title: str
    content_page_titles: list[str] = field(default_factory=list)


class TrainingDocumentForm:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.sections: list[EditableSection] = []
        self.selected_section_index: int | None = None

        self.title_var = tk.StringVar()
        self.section_title_var = tk.StringVar()
        self.content_title_var = tk.StringVar()
        self.validation_var = tk.StringVar(value="請輸入標題並至少新增一個章節。")
        self.template_path_var = tk.StringVar(value=str(DEFAULT_TEMPLATE_PATH))
        self.output_path_var = tk.StringVar()
        self.generation_status_var = tk.StringVar(value="請選擇輸出位置後產生 PPTX。")

        self._build()
        self._refresh()

    def set_document_title(self, title: str) -> None:
        self.title_var.set(title)
        self._refresh()

    def add_section(self, title: str) -> None:
        self.sections.append(EditableSection(title=title))
        self.selected_section_index = len(self.sections) - 1
        self._refresh()

    def rename_selected_section(self, title: str) -> None:
        if self.selected_section_index is None:
            return
        self.sections[self.selected_section_index].title = title
        self._refresh()

    def remove_selected_section(self) -> None:
        if self.selected_section_index is None:
            return
        del self.sections[self.selected_section_index]
        if not self.sections:
            self.selected_section_index = None
        else:
            self.selected_section_index = min(self.selected_section_index, len(self.sections) - 1)
        self._refresh()

    def move_selected_section_up(self) -> None:
        if self.selected_section_index is None or self.selected_section_index == 0:
            return
        index = self.selected_section_index
        self.sections[index - 1], self.sections[index] = self.sections[index], self.sections[index - 1]
        self.selected_section_index = index - 1
        self._refresh()

    def move_selected_section_down(self) -> None:
        if self.selected_section_index is None or self.selected_section_index >= len(self.sections) - 1:
            return
        index = self.selected_section_index
        self.sections[index + 1], self.sections[index] = self.sections[index], self.sections[index + 1]
        self.selected_section_index = index + 1
        self._refresh()

    def select_section(self, index: int) -> None:
        if index < 0 or index >= len(self.sections):
            self.selected_section_index = None
        else:
            self.selected_section_index = index
        self._refresh()

    def add_content_page(self, title: str) -> None:
        section = self._selected_section()
        if section is None:
            return
        section.content_page_titles.append(title)
        self._refresh()

    def rename_content_page(self, index: int, title: str) -> None:
        section = self._selected_section()
        if section is None or index < 0 or index >= len(section.content_page_titles):
            return
        section.content_page_titles[index] = title
        self._refresh()

    def remove_content_page(self, index: int) -> None:
        section = self._selected_section()
        if section is None or index < 0 or index >= len(section.content_page_titles):
            return
        del section.content_page_titles[index]
        self._refresh()

    def move_content_page_up(self, index: int) -> None:
        section = self._selected_section()
        if section is None or index <= 0 or index >= len(section.content_page_titles):
            return
        pages = section.content_page_titles
        pages[index - 1], pages[index] = pages[index], pages[index - 1]
        self._refresh()

    def move_content_page_down(self, index: int) -> None:
        section = self._selected_section()
        if section is None or index < 0 or index >= len(section.content_page_titles) - 1:
            return
        pages = section.content_page_titles
        pages[index + 1], pages[index] = pages[index], pages[index + 1]
        self._refresh()

    def current_document(self) -> TrainingDocumentInput:
        return TrainingDocumentInput(
            title=self.title_var.get(),
            sections=[
                SectionInput(title=section.title, content_page_titles=list(section.content_page_titles))
                for section in self.sections
            ],
        )

    def validation_messages(self) -> list[str]:
        return validate_training_document_input(self.current_document()).messages

    def preview_text(self) -> str:
        document = self.current_document()
        if validate_training_document_input(document).messages:
            return ""
        plan = generate_page_plan(document)
        labels = {
            "cover": "封面",
            "table_of_contents": "目錄",
            "section_start": "章節",
            "content": "內文",
        }
        lines = []
        for index, page in enumerate(plan.pages, start=1):
            label = labels[page.kind]
            if page.kind == "table_of_contents":
                lines.append(f"{index}. {label}")
            else:
                lines.append(f"{index}. {label} - {page.title}")
        return "\n".join(lines)

    def set_template_path(self, template_path: str | Path) -> None:
        self.template_path_var.set(str(template_path))

    def generate_to_path(self, output_path: Path, template_path: Path | None = None) -> Path | None:
        messages = self.validation_messages()
        if messages:
            self.generation_status_var.set("無法產生：" + " ".join(_localize_validation_messages(messages)))
            self._refresh()
            return None

        try:
            plan = generate_page_plan(self.current_document())
            selected_template_path = template_path or self._selected_template_path()
            path = export_page_plan_to_pptx(plan, output_path, template_path=selected_template_path)
        except Exception as error:
            self.generation_status_var.set(f"產生失敗：{error}")
            return None

        verification = verify_pptx_output(path, plan, template_path=selected_template_path)
        if not verification.is_valid:
            self.generation_status_var.set("驗證失敗：" + " ".join(verification.messages))
            return None

        self.output_path_var.set(str(path))
        self.generation_status_var.set(f"已產生並驗證：{path}")
        return path

    def _build(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        frame = ttk.Frame(self.root, padding=18, style="App.TFrame")
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=0, minsize=360)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(2, weight=1)

        ttk.Label(frame, text=APP_TITLE, style="Header.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        self.subtitle_label = ttk.Label(
            frame,
            text="建立教學文件架構後，再產生 PPTX 初稿。",
            style="Body.TLabel",
        )
        self.subtitle_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 14))

        left = ttk.Frame(frame, width=360, style="App.TFrame")
        self.left_panel = left
        left.grid(row=2, column=0, sticky="nsew", padx=(0, 14))
        left.grid_propagate(False)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(3, weight=1)
        left.rowconfigure(7, weight=1)

        right = ttk.Frame(frame, style="App.TFrame")
        right.grid(row=2, column=1, sticky="nsew", padx=(14, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        self.title_label = ttk.Label(left, text="PPT 標題", style="Body.TLabel")
        self.title_label.grid(row=0, column=0, sticky="w")
        ttk.Entry(left, textvariable=self.title_var).grid(row=1, column=0, sticky="ew", pady=(3, 12))
        self.title_var.trace_add("write", lambda *_: self._refresh())

        self.sections_label = ttk.Label(left, text="章節", style="Body.TLabel")
        self.sections_label.grid(row=2, column=0, sticky="w")
        self.section_list = tk.Listbox(left, height=8, exportselection=False)
        self.section_list.grid(row=3, column=0, sticky="nsew")
        self.section_list.bind("<<ListboxSelect>>", self._on_section_selected)

        section_controls = ttk.Frame(left, style="App.TFrame")
        section_controls.grid(row=4, column=0, sticky="ew", pady=(8, 12))
        section_controls.columnconfigure(0, weight=1)
        section_controls.columnconfigure(1, weight=1)
        ttk.Entry(section_controls, textvariable=self.section_title_var).grid(row=0, column=0, columnspan=2, sticky="ew")
        self.add_section_button = ttk.Button(
            section_controls,
            text="新增章節",
            command=self._add_section_from_entry,
        )
        self.add_section_button.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(section_controls, text="改名", command=self._rename_section_from_entry).grid(row=2, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(section_controls, text="刪除", command=self.remove_selected_section).grid(row=2, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))
        ttk.Button(section_controls, text="上移", command=self.move_selected_section_up).grid(row=3, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(section_controls, text="下移", command=self.move_selected_section_down).grid(row=3, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))

        self.content_pages_label = ttk.Label(left, text="選取章節的內文頁", style="Body.TLabel")
        self.content_pages_label.grid(row=5, column=0, sticky="w")
        self.content_list = tk.Listbox(left, height=7, exportselection=False)
        self.content_list.grid(row=7, column=0, sticky="nsew")

        content_controls = ttk.Frame(left, style="App.TFrame")
        content_controls.grid(row=8, column=0, sticky="ew", pady=(8, 0))
        content_controls.columnconfigure(0, weight=1)
        content_controls.columnconfigure(1, weight=1)
        ttk.Entry(content_controls, textvariable=self.content_title_var).grid(row=0, column=0, columnspan=2, sticky="ew")
        self.add_content_button = ttk.Button(
            content_controls,
            text="新增內文頁",
            command=self._add_content_from_entry,
        )
        self.add_content_button.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(content_controls, text="改名", command=self._rename_content_from_entry).grid(row=2, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(content_controls, text="刪除", command=self._remove_selected_content).grid(row=2, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))
        ttk.Button(content_controls, text="上移", command=self._move_selected_content_up).grid(row=3, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(content_controls, text="下移", command=self._move_selected_content_down).grid(row=3, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))

        self.preview_label = ttk.Label(right, text="投影片預覽", style="Body.TLabel")
        self.preview_label.grid(row=0, column=0, sticky="w")
        self.preview = tk.Text(right, height=20, wrap="word", borderwidth=1, relief="solid")
        self.preview.grid(row=1, column=0, sticky="nsew", pady=(3, 12))
        self.preview.configure(state="disabled")

        self.validation_label = ttk.Label(right, textvariable=self.validation_var, style="Status.TLabel", padding=(8, 6))
        self.validation_label.grid(row=2, column=0, sticky="ew")

        output = ttk.Frame(right, style="App.TFrame")
        output.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        output.columnconfigure(0, weight=1)
        self.template_label = ttk.Label(output, text="樣板 PPTX", style="Body.TLabel")
        self.template_label.grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Entry(output, textvariable=self.template_path_var).grid(row=1, column=0, sticky="ew", pady=(3, 10))
        self.choose_template_button = ttk.Button(output, text="選擇樣板", command=self._choose_template_path)
        self.choose_template_button.grid(row=1, column=1, padx=(6, 0), pady=(3, 10))

        ttk.Label(output, text="輸出位置", style="Body.TLabel").grid(row=2, column=0, columnspan=3, sticky="w")
        ttk.Entry(output, textvariable=self.output_path_var).grid(row=3, column=0, sticky="ew", pady=(3, 0))
        ttk.Button(output, text="選擇位置", command=self._choose_output_path).grid(row=3, column=1, padx=(6, 0), pady=(3, 0))
        self.generate_button = ttk.Button(output, text="產生 PPTX", command=self._generate_from_output_path)
        self.generate_button.grid(
            row=3,
            column=2,
            padx=(6, 0),
            pady=(3, 0),
        )

        ttk.Label(right, textvariable=self.generation_status_var, style="Status.TLabel", padding=(8, 6)).grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(8, 0),
        )

    def _refresh(self) -> None:
        if hasattr(self, "section_list"):
            self._refresh_sections()
            self._refresh_content_pages()
            self._refresh_preview_and_validation()

    def _refresh_sections(self) -> None:
        self.section_list.delete(0, tk.END)
        for index, section in enumerate(self.sections, start=1):
            self.section_list.insert(tk.END, f"{index}. {section.title}")
        if self.selected_section_index is not None and self.sections:
            self.section_list.selection_set(self.selected_section_index)
            self.section_list.activate(self.selected_section_index)

    def _refresh_content_pages(self) -> None:
        self.content_list.delete(0, tk.END)
        section = self._selected_section()
        if section is None:
            return
        for index, title in enumerate(section.content_page_titles, start=1):
            self.content_list.insert(tk.END, f"{index}. {title}")

    def _refresh_preview_and_validation(self) -> None:
        messages = self.validation_messages()
        self.validation_var.set("就緒" if not messages else " ".join(_localize_validation_messages(messages)))
        self.preview.configure(state="normal")
        self.preview.delete("1.0", tk.END)
        preview = self.preview_text()
        if preview:
            self.preview.insert("1.0", preview)
        self.preview.configure(state="disabled")

    def _selected_section(self) -> EditableSection | None:
        if self.selected_section_index is None:
            return None
        if self.selected_section_index < 0 or self.selected_section_index >= len(self.sections):
            return None
        return self.sections[self.selected_section_index]

    def _on_section_selected(self, _event=None) -> None:
        selection = self.section_list.curselection()
        self.selected_section_index = selection[0] if selection else None
        self._refresh()

    def _add_section_from_entry(self) -> None:
        self.add_section(self.section_title_var.get())
        self.section_title_var.set("")

    def _rename_section_from_entry(self) -> None:
        self.rename_selected_section(self.section_title_var.get())

    def _add_content_from_entry(self) -> None:
        self.add_content_page(self.content_title_var.get())
        self.content_title_var.set("")

    def _rename_content_from_entry(self) -> None:
        index = self._selected_content_index()
        if index is not None:
            self.rename_content_page(index, self.content_title_var.get())

    def _remove_selected_content(self) -> None:
        index = self._selected_content_index()
        if index is not None:
            self.remove_content_page(index)

    def _move_selected_content_up(self) -> None:
        index = self._selected_content_index()
        if index is not None:
            self.move_content_page_up(index)

    def _move_selected_content_down(self) -> None:
        index = self._selected_content_index()
        if index is not None:
            self.move_content_page_down(index)

    def _selected_content_index(self) -> int | None:
        selection = self.content_list.curselection()
        return selection[0] if selection else None

    def _selected_template_path(self) -> Path:
        raw_path = self.template_path_var.get().strip()
        return Path(raw_path) if raw_path else DEFAULT_TEMPLATE_PATH

    def _choose_template_path(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="選擇樣板 PPTX",
            filetypes=[("PowerPoint files", "*.pptx")],
        )
        if filename:
            self.template_path_var.set(filename)

    def _choose_output_path(self) -> None:
        filename = filedialog.asksaveasfilename(
            parent=self.root,
            title="儲存產生的 PPTX",
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx")],
        )
        if filename:
            self.output_path_var.set(filename)

    def _generate_from_output_path(self) -> None:
        raw_path = self.output_path_var.get().strip()
        if not raw_path:
            self.generation_status_var.set("請先選擇輸出位置。")
            return
        self.generate_to_path(Path(raw_path))


def create_main_window() -> tk.Tk:
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry(APP_GEOMETRY)
    root.minsize(840, 620)

    _configure_styles(root)
    root.form = TrainingDocumentForm(root)

    return root


def _configure_styles(root: tk.Tk) -> None:
    style = ttk.Style(root)
    style.configure("App.TFrame", background="#f5f6f8")
    style.configure("Header.TLabel", background="#f5f6f8", font=("Segoe UI", 18, "bold"))
    style.configure("Body.TLabel", background="#f5f6f8", font=("Segoe UI", 10))
    style.configure("Status.TLabel", background="#eef1f5", font=("Segoe UI", 9))


def _localize_validation_messages(messages: list[str]) -> list[str]:
    return [_localize_validation_message(message) for message in messages]


def _localize_validation_message(message: str) -> str:
    if message == "Document title is required.":
        return "請輸入 PPT 標題。"
    if message == "At least one section is required.":
        return "請至少新增一個章節。"

    match = re.fullmatch(r"Section (\d+) title is required\.", message)
    if match:
        return f"第 {match.group(1)} 個章節需要標題。"

    match = re.fullmatch(r"Section (\d+) must include at least one content page\.", message)
    if match:
        return f"第 {match.group(1)} 個章節至少需要一個內文頁。"

    match = re.fullmatch(r"Section (\d+) content page (\d+) title is required\.", message)
    if match:
        return f"第 {match.group(1)} 個章節的第 {match.group(2)} 個內文頁需要標題。"

    return message
