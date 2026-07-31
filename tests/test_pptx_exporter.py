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
from pptx_quicktool.pptx_exporter import DEFAULT_TEMPLATE_PATH, export_page_plan_to_pptx
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
                ["Setup", "回主目錄"],
                ["Install Tools", "回主目錄"],
                ["Connect Controller", "回主目錄"],
                ["Operation", "回主目錄"],
                ["Open Project", "回主目錄"],
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
            back_link = shape_with_text(presentation.slides[slide_index], "回主目錄")
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

    def test_exports_reference_size_and_consistent_title_hierarchy(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(title="Setup", content_page_titles=["Install Tools"]),
                SectionInput(title="Operation", content_page_titles=["Open Project"]),
            ],
        )
        plan = generate_page_plan(document)
        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "training-document-format.pptx"
        export_page_plan_to_pptx(plan, output_path)

        presentation = Presentation(output_path)
        self.assertEqual(presentation.slide_width, 12192000)
        self.assertEqual(presentation.slide_height, 6858000)

        title_shapes = [
            shape_with_text(presentation.slides[0], "TwinCAT Training"),
            shape_with_text(presentation.slides[1], "\u76ee\u9304"),
            shape_with_text(presentation.slides[2], "Setup"),
            shape_with_text(presentation.slides[3], "Install Tools"),
        ]
        self.assertEqual({shape.left for shape in title_shapes}, {685800})
        self.assertEqual({shape.top for shape in title_shapes}, {457200})
        self.assertEqual({shape.text_frame.paragraphs[0].font.size.pt for shape in title_shapes}, {34.0})

    def test_exports_with_default_training_document_template_layouts(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[SectionInput(title="Setup", content_page_titles=["Install Tools"])],
        )
        plan = generate_page_plan(document)
        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "training-document-default-template.pptx"

        export_page_plan_to_pptx(plan, output_path)

        presentation = Presentation(output_path)
        self.assertTrue(DEFAULT_TEMPLATE_PATH.exists())
        self.assertEqual(
            [slide.slide_layout.name for slide in presentation.slides],
            ["Title with picture", "Contents", "Contents", "Text"],
        )

    def test_exports_with_specified_template_path_without_modifying_source(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[SectionInput(title="Setup", content_page_titles=["Install Tools"])],
        )
        plan = generate_page_plan(document)
        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "training-document-specified-template.pptx"
        template_mtime = DEFAULT_TEMPLATE_PATH.stat().st_mtime_ns

        export_page_plan_to_pptx(plan, output_path, template_path=DEFAULT_TEMPLATE_PATH)

        presentation = Presentation(output_path)
        self.assertEqual(len(presentation.slides), 4)
        self.assertEqual(presentation.slides[0].slide_layout.name, "Title with picture")
        self.assertEqual(DEFAULT_TEMPLATE_PATH.stat().st_mtime_ns, template_mtime)


if __name__ == "__main__":
    unittest.main()
