from __future__ import annotations

from dataclasses import dataclass

from .training_document import TrainingDocumentInput


@dataclass(frozen=True)
class PlannedPage:
    page_id: str
    kind: str
    title: str


@dataclass(frozen=True)
class TableOfContentsEntry:
    title: str
    target_page_id: str


@dataclass(frozen=True)
class PagePlan:
    pages: list[PlannedPage]
    table_of_contents_entries: list[TableOfContentsEntry]


def generate_page_plan(document: TrainingDocumentInput) -> PagePlan:
    pages = [
        PlannedPage(page_id="cover", kind="cover", title=document.title),
        PlannedPage(page_id="table-of-contents", kind="table_of_contents", title="\u76ee\u9304"),
    ]
    table_of_contents_entries: list[TableOfContentsEntry] = []

    for section_index, section in enumerate(document.sections, start=1):
        section_page_id = f"section-{section_index}"
        pages.append(
            PlannedPage(
                page_id=section_page_id,
                kind="section_start",
                title=section.title,
            )
        )
        table_of_contents_entries.append(
            TableOfContentsEntry(title=section.title, target_page_id=section_page_id)
        )
        for content_index, content_page_title in enumerate(section.content_page_titles, start=1):
            pages.append(
                PlannedPage(
                    page_id=f"section-{section_index}-content-{content_index}",
                    kind="content",
                    title=content_page_title,
                )
            )

    return PagePlan(pages=pages, table_of_contents_entries=table_of_contents_entries)
