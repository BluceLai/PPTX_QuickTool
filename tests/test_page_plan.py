from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx_quicktool.page_plan import generate_page_plan
from pptx_quicktool.training_document import SectionInput, TrainingDocumentInput


class PagePlanTests(unittest.TestCase):
    def test_minimal_training_document_produces_core_pages(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(title="Setup", content_page_titles=["Install Tools"]),
            ],
        )

        plan = generate_page_plan(document)

        self.assertEqual(
            [(page.kind, page.title) for page in plan.pages],
            [
                ("cover", "TwinCAT Training"),
                ("table_of_contents", "\u76ee\u9304"),
                ("section_start", "Setup"),
                ("content", "Install Tools"),
            ],
        )

    def test_page_order_matches_entered_section_and_content_order(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(
                    title="Setup",
                    content_page_titles=["Install Tools", "Connect Controller"],
                ),
                SectionInput(
                    title="Operation",
                    content_page_titles=["Open Project"],
                ),
            ],
        )

        plan = generate_page_plan(document)

        self.assertEqual(
            [(page.kind, page.title) for page in plan.pages],
            [
                ("cover", "TwinCAT Training"),
                ("table_of_contents", "\u76ee\u9304"),
                ("section_start", "Setup"),
                ("content", "Install Tools"),
                ("content", "Connect Controller"),
                ("section_start", "Operation"),
                ("content", "Open Project"),
            ],
        )

    def test_page_plan_exposes_stable_identifiers_for_navigation(self) -> None:
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

        self.assertEqual(
            [page.page_id for page in plan.pages],
            [
                "cover",
                "table-of-contents",
                "section-1",
                "section-1-content-1",
                "section-1-content-2",
                "section-2",
                "section-2-content-1",
            ],
        )
        self.assertEqual(
            [(entry.title, entry.target_page_id) for entry in plan.table_of_contents_entries],
            [
                ("Setup", "section-1"),
                ("Operation", "section-2"),
            ],
        )

    def test_section_without_content_pages_gets_a_blank_content_page_named_after_the_section(self) -> None:
        document = TrainingDocumentInput(
            title="TwinCAT Training",
            sections=[
                SectionInput(title="Foreword", content_page_titles=[]),
            ],
        )

        plan = generate_page_plan(document)

        self.assertEqual(
            [(page.page_id, page.kind, page.title) for page in plan.pages],
            [
                ("cover", "cover", "TwinCAT Training"),
                ("table-of-contents", "table_of_contents", "\u76ee\u9304"),
                ("section-1", "section_start", "Foreword"),
                ("section-1-content-1", "content", "Foreword"),
            ],
        )


if __name__ == "__main__":
    unittest.main()
