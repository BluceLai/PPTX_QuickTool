from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn

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


def slide_combined_text(slide) -> str:
    texts = []
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip():
            texts.append(shape.text.strip().replace("\v", "\n"))
    return "\n".join(texts)


def linked_slide_indices(presentation: Presentation, slide) -> list[int]:
    indices = []
    for shape in slide.shapes:
        target_slide = shape.click_action.target_slide
        if target_slide is not None:
            indices.append(slide_index_for_target(presentation, target_slide))
        for hyperlink in shape._element.xpath(".//a:rPr/a:hlinkClick"):
            relationship_id = hyperlink.get(qn("r:id"))
            if not relationship_id:
                continue
            relationship = slide.part.rels[relationship_id]
            target_part = relationship.target_part
            indices.append(slide_part_index(presentation, target_part))
    return indices


def slide_part_index(presentation: Presentation, target_part) -> int:
    for index, candidate_slide in enumerate(presentation.slides):
        if candidate_slide.part == target_part:
            return index
    raise AssertionError("Target slide part was not found in presentation")


def placeholder_shape_with_text(slide, text: str):
    shape = shape_with_text(slide, text)
    if not shape.is_placeholder:
        raise AssertionError(f"Shape with text is not a placeholder: {text}")
    return shape


def title_placeholder(layout):
    for placeholder in layout.placeholders:
        if placeholder.placeholder_format.type == PP_PLACEHOLDER.TITLE:
            return placeholder
    raise AssertionError("Layout has no title placeholder")


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
        self.assertIn("TwinCAT Training", slide_combined_text(presentation.slides[0]))
        self.assertIn("目錄", slide_combined_text(presentation.slides[1]))
        self.assertIn("Setup", slide_combined_text(presentation.slides[1]))
        self.assertIn("Install Tools", slide_combined_text(presentation.slides[1]))
        self.assertIn("Connect Controller", slide_combined_text(presentation.slides[1]))
        self.assertIn("Operation", slide_combined_text(presentation.slides[1]))
        self.assertIn("Setup", slide_combined_text(presentation.slides[2]))
        self.assertIn("Install Tools", slide_combined_text(presentation.slides[3]))
        self.assertNotIn("回主目錄", slide_combined_text(presentation.slides[3]))

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
        self.assertEqual(linked_slide_indices(presentation, presentation.slides[1]), [2, 5])
        self.assertEqual(linked_slide_indices(presentation, presentation.slides[2]), [1])
        self.assertIn("目錄", slide_combined_text(presentation.slides[2]))
        self.assertIn("Operation", slide_combined_text(presentation.slides[2]))
        self.assertEqual(linked_slide_indices(presentation, presentation.slides[5]), [1])
        for slide_index in [3, 4, 6]:
            self.assertEqual(linked_slide_indices(presentation, presentation.slides[slide_index]), [])

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
        self.assertEqual(linked_slide_indices(presentation, presentation.slides[1]), [2, 4])

    def test_agenda_text_leaves_bullets_and_numbering_to_the_template(self) -> None:
        document = TrainingDocumentInput(
            title="測試文件",
            sections=[
                SectionInput(title="測試1", content_page_titles=["1", "2"]),
                SectionInput(title="測試2", content_page_titles=["1", "2"]),
            ],
        )
        plan = generate_page_plan(document)
        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "training-document-template-agenda-text.pptx"

        export_page_plan_to_pptx(plan, output_path)

        presentation = Presentation(output_path)
        table_of_contents_text = slide_combined_text(presentation.slides[1])
        section_text = slide_combined_text(presentation.slides[2])

        self.assertIn("測試1", table_of_contents_text)
        self.assertIn("測試2", table_of_contents_text)
        self.assertIn("1", table_of_contents_text)
        self.assertNotIn("1. 測試1", table_of_contents_text)
        self.assertNotIn("2. 測試2", table_of_contents_text)
        self.assertNotIn("- 1", table_of_contents_text)
        self.assertNotIn("- 2", table_of_contents_text)
        self.assertNotIn("1. 測試1", section_text)
        self.assertNotIn("- 1", section_text)

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
            placeholder_shape_with_text(presentation.slides[0], "TwinCAT Training"),
            placeholder_shape_with_text(presentation.slides[1], "\u76ee\u9304"),
            placeholder_shape_with_text(presentation.slides[2], "Setup"),
            placeholder_shape_with_text(presentation.slides[3], "Install Tools"),
        ]
        template = Presentation(DEFAULT_TEMPLATE_PATH)
        template_title_placeholder = title_placeholder(template.slide_layouts[0])
        self.assertEqual({shape.left for shape in title_shapes}, {template_title_placeholder.left})
        self.assertEqual({shape.top for shape in title_shapes}, {template_title_placeholder.top})
        self.assertEqual({shape.text_frame.paragraphs[0].runs[0].font.name for shape in title_shapes}, {None})
        self.assertEqual({shape.text_frame.paragraphs[0].runs[0].font.size for shape in title_shapes}, {None})

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
