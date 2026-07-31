from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx import Presentation

from pptx_quicktool.page_plan import generate_page_plan
from pptx_quicktool.pptx_exporter import export_page_plan_to_pptx
from pptx_quicktool.training_document import SectionInput, TrainingDocumentInput


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


def shape_with_text(slide, text: str):
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip().replace("\v", "\n") == text:
            return shape
    raise AssertionError(f"Could not find shape with text: {text}")


def slide_index_for_target(presentation: Presentation, target_slide) -> int:
    for index, slide in enumerate(presentation.slides):
        if slide == target_slide:
            return index
    raise AssertionError("Target slide was not found in presentation")


class PptxExporterTests(unittest.TestCase):
    def test_exports_valid_pptx_with_expected_slides_and_text(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(
                    title="Setup",
                    content_page_titles=["Install Tools", "Connect Controller"],
                ),
                SectionInput(title="Operation", content_page_titles=["Open Project"]),
            ],
        )
        plan = generate_page_plan(document)

        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "training-document.pptx"
        if output_path.exists():
            output_path.unlink()

        export_page_plan_to_pptx(plan, output_path)

        self.assertTrue(output_path.exists())
        presentation = Presentation(output_path)
        self.assertEqual(len(presentation.slides), 7)
        self.assertEqual(
            slide_texts(output_path),
            [
                ["TwinCAT Training"],
                ["\u76ee\u9304", "1. Setup", "2. Operation"],
                ["Setup", "Back to contents"],
                ["Install Tools", "Back to contents"],
                ["Connect Controller", "Back to contents"],
                ["Operation", "Back to contents"],
                ["Open Project", "Back to contents"],
            ],
        )

    def test_exports_table_of_contents_and_return_navigation_links(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(
                    title="Setup",
                    content_page_titles=["Install Tools", "Connect Controller"],
                ),
                SectionInput(title="Operation", content_page_titles=["Open Project"]),
            ],
        )
        plan = generate_page_plan(document)
        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "training-document-links.pptx"
        if output_path.exists():
            output_path.unlink()

        export_page_plan_to_pptx(plan, output_path)

        presentation = Presentation(output_path)
        setup_link = shape_with_text(presentation.slides[1], "1. Setup")
        operation_link = shape_with_text(presentation.slides[1], "2. Operation")
        self.assertEqual(
            slide_index_for_target(presentation, setup_link.click_action.target_slide),
            2,
        )
        self.assertEqual(
            slide_index_for_target(
                presentation,
                operation_link.click_action.target_slide,
            ),
            5,
        )
        for slide_index in [2, 3, 4, 5, 6]:
            back_link = shape_with_text(presentation.slides[slide_index], "Back to contents")
            self.assertEqual(
                slide_index_for_target(presentation, back_link.click_action.target_slide),
                1,
            )

    def test_table_of_contents_links_follow_reordered_sections(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(title="Operation", content_page_titles=["Open Project"]),
                SectionInput(
                    title="Setup",
                    content_page_titles=["Install Tools", "Connect Controller"],
                ),
            ],
        )
        plan = generate_page_plan(document)
        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "training-document-reordered-links.pptx"
        if output_path.exists():
            output_path.unlink()

        export_page_plan_to_pptx(plan, output_path)

        presentation = Presentation(output_path)
        operation_link = shape_with_text(presentation.slides[1], "1. Operation")
        setup_link = shape_with_text(presentation.slides[1], "2. Setup")
        self.assertEqual(
            slide_index_for_target(
                presentation,
                operation_link.click_action.target_slide,
            ),
            2,
        )
        self.assertEqual(
            slide_index_for_target(presentation, setup_link.click_action.target_slide),
            4,
        )


if __name__ == "__main__":
    unittest.main()
