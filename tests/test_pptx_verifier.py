from __future__ import annotations

import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"

if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from pptx_quicktool.page_plan import generate_page_plan
from pptx_quicktool.pptx_exporter import export_page_plan_to_pptx
from pptx_quicktool.pptx_verifier import verify_pptx_output
from pptx_quicktool.training_document import SectionInput, TrainingDocumentInput


class PptxVerifierTests(unittest.TestCase):
    def test_generated_pptx_passes_output_verification(self) -> None:
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
        output_path = output_dir / "verified-training-document.pptx"
        if output_path.exists():
            output_path.unlink()
        export_page_plan_to_pptx(plan, output_path)

        result = verify_pptx_output(output_path, plan)

        self.assertTrue(result.is_valid)
        self.assertEqual(result.messages, [])

    def test_verification_reports_open_failures(self) -> None:
        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "not-a-presentation.pptx"
        output_path.write_text("not a pptx", encoding="utf-8")
        plan = generate_page_plan(
            TrainingDocumentInput(
                title="TwinCAT Training",
                sections=[SectionInput(title="Setup", content_page_titles=["Install Tools"])],
            )
        )

        result = verify_pptx_output(output_path, plan)

        self.assertFalse(result.is_valid)
        self.assertEqual(result.messages, ["Output file cannot be opened as a PPTX."])

    def test_verification_reports_slide_count_mismatch(self) -> None:
        generated_plan = generate_page_plan(
            TrainingDocumentInput(
                title="TwinCAT Training",
                sections=[SectionInput(title="Setup", content_page_titles=["Install Tools"])],
            )
        )
        expected_plan = generate_page_plan(
            TrainingDocumentInput(
                title="TwinCAT Training",
                sections=[
                    SectionInput(title="Setup", content_page_titles=["Install Tools"]),
                    SectionInput(title="Operation", content_page_titles=["Open Project"]),
                ],
            )
        )
        output_dir = ROOT / ".tmp" / "test-output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "wrong-slide-count.pptx"
        export_page_plan_to_pptx(generated_plan, output_path)

        result = verify_pptx_output(output_path, expected_plan)

        self.assertFalse(result.is_valid)
        self.assertIn("Expected 6 slides, found 4.", result.messages)


if __name__ == "__main__":
    unittest.main()
