from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt

from .page_plan import PagePlan, PlannedPage


SLIDE_WIDTH = 12192000
SLIDE_HEIGHT = 6858000
DEFAULT_TEMPLATE_PATH = Path(__file__).resolve().parent / "assets" / "default-training-document-template.pptx"
RETURN_TO_CONTENTS_TEXT = "回主目錄"

LAYOUT_NAMES_BY_PAGE_KIND = {
    "cover": ("Title with picture",),
    "table_of_contents": ("Contents", "Text"),
    "section_start": ("Contents", "Text"),
    "content": ("Text",),
}


def export_page_plan_to_pptx(plan: PagePlan, output_path: Path, template_path: Path | None = None) -> Path:
    presentation = _new_presentation_from_template(template_path)
    slide_by_page_id = {}
    slide_pages = []
    for page in plan.pages:
        slide = presentation.slides.add_slide(_layout_for_page(presentation, page))
        _remove_slide_placeholders(slide)
        slide_by_page_id[page.page_id] = slide
        slide_pages.append((slide, page))

    table_of_contents_slide = slide_by_page_id["table-of-contents"]
    for slide, page in slide_pages:
        _populate_slide(slide, page, plan, slide_by_page_id, table_of_contents_slide)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(output_path)
    return output_path


def _new_presentation_from_template(template_path: Path | None):
    path = Path(template_path) if template_path is not None else DEFAULT_TEMPLATE_PATH
    if not path.exists():
        raise FileNotFoundError(f"Template PPTX was not found: {path}")

    presentation = Presentation(path)
    _remove_all_slides(presentation)
    return presentation


def _remove_all_slides(presentation) -> None:
    slide_id_list = presentation.slides._sldIdLst
    for slide_id in list(slide_id_list):
        presentation.part.drop_rel(slide_id.rId)
        slide_id_list.remove(slide_id)


def _layout_for_page(presentation, page: PlannedPage):
    wanted_names = LAYOUT_NAMES_BY_PAGE_KIND[page.kind]
    for name in wanted_names:
        for layout in presentation.slide_layouts:
            if layout.name == name:
                return layout
    return presentation.slide_layouts[0]


def _remove_slide_placeholders(slide) -> None:
    for shape in list(slide.shapes):
        if shape.is_placeholder:
            shape._element.getparent().remove(shape._element)


def _populate_slide(slide, page: PlannedPage, plan: PagePlan, slide_by_page_id, table_of_contents_slide) -> None:
    _add_title(slide, page.title)
    if page.kind == "table_of_contents":
        _add_table_of_contents_links(slide, plan, slide_by_page_id)
    elif page.kind in {"section_start", "content"}:
        _add_return_link(slide, table_of_contents_slide)


def _add_title(slide, text: str) -> None:
    box = slide.shapes.add_textbox(Inches(0.75), Inches(0.5), Inches(11.8), Inches(0.75))
    text_frame = box.text_frame
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.name = "Segoe UI"
    paragraph.font.size = Pt(34)
    paragraph.font.bold = True


def _add_link_text(slide, text: str, left, top, width, height):
    box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = box.text_frame
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.name = "Segoe UI"
    paragraph.font.size = Pt(22)
    return box


def _add_table_of_contents_links(slide, plan: PagePlan, slide_by_page_id) -> None:
    for entry_index, entry in enumerate(plan.table_of_contents_entries, start=1):
        link = _add_link_text(
            slide,
            f"{entry_index}. {entry.title}",
            Inches(1.15),
            Inches(1.65 + (entry_index - 1) * 0.48),
            Inches(10.4),
            Inches(0.4),
        )
        link.click_action.target_slide = slide_by_page_id[entry.target_page_id]


def _add_return_link(slide, table_of_contents_slide) -> None:
    link = _add_link_text(
        slide,
        RETURN_TO_CONTENTS_TEXT,
        Inches(10.0),
        Inches(6.55),
        Inches(2.3),
        Inches(0.35),
    )
    link.text_frame.paragraphs[0].font.size = Pt(12)
    link.click_action.target_slide = table_of_contents_slide
