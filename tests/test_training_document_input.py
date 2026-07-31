from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx_quicktool.training_document import (
    SectionInput,
    TrainingDocumentInput,
    validate_training_document_input,
)


class TrainingDocumentInputTests(unittest.TestCase):
    def test_training_document_keeps_ordered_sections_and_pages(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(
                    title="Setup",
                    content_page_titles=["Install Tools", "Connect Controller"],
                ),
                SectionInput(
                    title="Operation",
                    content_page_titles=["Open Project", "Run Check"],
                ),
            ],
        )

        self.assertEqual(document.title, "TwinCAT Training")
        self.assertEqual([section.title for section in document.sections], ["Setup", "Operation"])
        self.assertEqual(document.sections[0].content_page_titles, ["Install Tools", "Connect Controller"])
        self.assertEqual(document.sections[1].content_page_titles, ["Open Project", "Run Check"])

    def test_valid_training_document_has_no_validation_errors(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(title="Setup", content_page_titles=["Install Tools"]),
            ],
        )

        result = validate_training_document_input(document)

        self.assertTrue(result.is_valid)
        self.assertEqual(result.messages, [])

    def test_validation_rejects_empty_document_title(self) -> None:
        document = TrainingDocumentInput(
            title="  ",
            sections=[
                SectionInput(title="Setup", content_page_titles=["Install Tools"]),
            ],
        )

        result = validate_training_document_input(document)

        self.assertFalse(result.is_valid)
        self.assertEqual(result.messages, ["Document title is required."])

    def test_validation_rejects_missing_sections(self) -> None:
        document = TrainingDocumentInput(title="TwinCAT Training", sections=[])

        result = validate_training_document_input(document)

        self.assertFalse(result.is_valid)
        self.assertEqual(result.messages, ["At least one section is required."])

    def test_validation_rejects_empty_section_titles(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(title="Setup", content_page_titles=["Install Tools"]),
                SectionInput(title=" ", content_page_titles=["Open Project"]),
            ],
        )

        result = validate_training_document_input(document)

        self.assertFalse(result.is_valid)
        self.assertEqual(result.messages, ["Section 2 title is required."])

    def test_validation_rejects_sections_without_content_pages(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(title="Setup", content_page_titles=[]),
            ],
        )

        result = validate_training_document_input(document)

        self.assertFalse(result.is_valid)
        self.assertEqual(result.messages, ["Section 1 must include at least one content page."])

    def test_validation_rejects_empty_content_page_titles(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(title="Setup", content_page_titles=["Install Tools", " "]),
            ],
        )

        result = validate_training_document_input(document)

        self.assertFalse(result.is_valid)
        self.assertEqual(result.messages, ["Section 1 content page 2 title is required."])


if __name__ == "__main__":
    unittest.main()
