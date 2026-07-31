from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt

from .page_plan import PagePlan


SLIDE_WIDTH = 12192000
SLIDE_HEIGHT = 6858000


def export_page_plan_to_pptx(plan: PagePlan, output_path: Path) -> Path:
    presentation = Presentation()
    presentation.slide_width = SLIDE_WIDTH
    presentation.slide_height = SLIDE_HEIGHT

    for page in plan.pages:
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        if page.kind == "table_of_contents":
            _add_title(slide, page.title)
            _add_body(slide, _format_table_of_contents(plan))
        else:
            _add_title(slide, page.title)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(output_path)
    return output_path


def _add_title(slide, text: str) -> None:
    box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.8), Inches(0.75))
    text_frame = box.text_frame
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.name = "Segoe UI"
    paragraph.font.size = Pt(32)
    paragraph.font.bold = True


def _add_body(slide, text: str) -> None:
    box = slide.shapes.add_textbox(Inches(1.15), Inches(1.75), Inches(10.4), Inches(4.8))
    text_frame = box.text_frame
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.name = "Segoe UI"
    paragraph.font.size = Pt(22)


def _format_table_of_contents(plan: PagePlan) -> str:
    return "\n".join(
        f"{entry_index}. {entry.title}"
        for entry_index, entry in enumerate(plan.table_of_contents_entries, start=1)
    )
