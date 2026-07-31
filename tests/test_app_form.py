from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx_quicktool.app import create_main_window
from pptx import Presentation


def slide_texts(path: Path) -> list[list[str]]:
    presentation = Presentation(path)
    slides: list[list[str]] = []
    for slide in presentation.slides:
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text.strip().replace("\v", "\n"))
        slides.append(texts)
    return slides


class AppFormTests(unittest.TestCase):
    def test_form_edits_document_structure_and_preview(self) -> None:
        root = create_main_window()
        try:
            form = root.form

            form.set_document_title("TwinCAT Training")
            form.add_section("Setup")
            form.add_content_page("Install Tools")
            form.add_content_page("Connect Controller")
            form.add_section("Operation")
            form.add_content_page("Open Project")

            document = form.current_document()
            self.assertEqual(document.title, "TwinCAT Training")
            self.assertEqual([section.title for section in document.sections], ["Setup", "Operation"])
            self.assertEqual(document.sections[0].content_page_titles, ["Install Tools", "Connect Controller"])
            self.assertEqual(document.sections[1].content_page_titles, ["Open Project"])
            self.assertEqual(form.validation_messages(), [])
            self.assertEqual(
                form.preview_text(),
                (
                    "1. Cover - TwinCAT Training\n"
                    "2. Table of contents\n"
                    "3. Section - Setup\n"
                    "4. Content - Install Tools\n"
                    "5. Content - Connect Controller\n"
                    "6. Section - Operation\n"
                    "7. Content - Open Project"
                ),
            )
        finally:
            root.update()
            root.destroy()

    def test_form_renames_removes_and_reorders_sections_and_pages(self) -> None:
        root = create_main_window()
        try:
            form = root.form

            form.set_document_title("TwinCAT Training")
            form.add_section("Setup")
            form.add_content_page("Install Tools")
            form.add_content_page("Connect Controller")
            form.add_section("Operation")
            form.add_content_page("Open Project")

            form.select_section(1)
            form.rename_selected_section("Operate")
            form.move_selected_section_up()
            form.select_section(0)
            form.rename_content_page(0, "Open Existing Project")
            form.add_content_page("Run Check")
            form.move_content_page_down(0)
            form.remove_content_page(0)

            document = form.current_document()
            self.assertEqual([section.title for section in document.sections], ["Operate", "Setup"])
            self.assertEqual(document.sections[0].content_page_titles, ["Open Existing Project"])
            self.assertEqual(document.sections[1].content_page_titles, ["Install Tools", "Connect Controller"])
        finally:
            root.update()
            root.destroy()

    def test_form_shows_validation_messages_from_input_model(self) -> None:
        root = create_main_window()
        try:
            form = root.form

            form.set_document_title("")
            form.add_section("")

            self.assertEqual(
                form.validation_messages(),
                [
                    "Document title is required.",
                    "Section 1 title is required.",
                    "Section 1 must include at least one content page.",
                ],
            )
            self.assertEqual(
                form.validation_var.get(),
                (
                    "Document title is required. "
                    "Section 1 title is required. "
                    "Section 1 must include at least one content page."
                ),
            )
            self.assertEqual(form.preview_text(), "")
        finally:
            root.update()
            root.destroy()

    def test_form_generates_pptx_from_current_data(self) -> None:
        root = create_main_window()
        try:
            form = root.form
            form.set_document_title("TwinCAT Training")
            form.add_section("Setup")
            form.add_content_page("Install Tools")
            form.add_section("Operation")
            form.add_content_page("Open Project")

            output_dir = ROOT / ".tmp" / "test-output"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / "ui-generated.pptx"
            if output_path.exists():
                output_path.unlink()

            result = form.generate_to_path(output_path)

            self.assertEqual(result, output_path)
            self.assertTrue(output_path.exists())
            self.assertEqual(
                slide_texts(output_path),
                [
                    ["TwinCAT Training"],
                    ["\u76ee\u9304", "1. Setup", "2. Operation"],
                    ["Setup", "Back to contents"],
                    ["Install Tools", "Back to contents"],
                    ["Operation", "Back to contents"],
                    ["Open Project", "Back to contents"],
                ],
            )
            self.assertEqual(form.generation_status_var.get(), f"Generated and verified: {output_path}")
        finally:
            root.update()
            root.destroy()

    def test_form_blocks_generation_when_validation_fails(self) -> None:
        root = create_main_window()
        try:
            form = root.form
            output_dir = ROOT / ".tmp" / "test-output"
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / "invalid-ui-generated.pptx"
            if output_path.exists():
                output_path.unlink()

            result = form.generate_to_path(output_path)

            self.assertIsNone(result)
            self.assertFalse(output_path.exists())
            self.assertIn("Document title is required.", form.generation_status_var.get())
        finally:
            root.update()
            root.destroy()

    def test_form_reports_generation_errors(self) -> None:
        root = create_main_window()
        try:
            form = root.form
            form.set_document_title("TwinCAT Training")
            form.add_section("Setup")
            form.add_content_page("Install Tools")
            output_dir = ROOT / ".tmp" / "test-output"
            output_dir.mkdir(parents=True, exist_ok=True)

            result = form.generate_to_path(output_dir)

            self.assertIsNone(result)
            self.assertTrue(form.generation_status_var.get().startswith("Generation failed:"))
        finally:
            root.update()
            root.destroy()


if __name__ == "__main__":
    unittest.main()
