from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SectionInput:
    title: str
    content_page_titles: list[str]


@dataclass(frozen=True)
class TrainingDocumentInput:
    title: str
    sections: list[SectionInput]


@dataclass(frozen=True)
class ValidationResult:
    messages: list[str]

    @property
    def is_valid(self) -> bool:
        return len(self.messages) == 0


def validate_training_document_input(document: TrainingDocumentInput) -> ValidationResult:
    messages: list[str] = []
    if not document.title.strip():
        messages.append("Document title is required.")
    if len(document.sections) == 0:
        messages.append("At least one section is required.")
    for section_index, section in enumerate(document.sections, start=1):
        if not section.title.strip():
            messages.append(f"Section {section_index} title is required.")
        if len(section.content_page_titles) == 0:
            messages.append(f"Section {section_index} must include at least one content page.")
        seen_page_titles: dict[str, str] = {}
        for page_index, page_title in enumerate(section.content_page_titles, start=1):
            normalized_title = page_title.strip()
            if not normalized_title:
                messages.append(f"Section {section_index} content page {page_index} title is required.")
                continue
            duplicate_key = normalized_title.casefold()
            if duplicate_key in seen_page_titles:
                messages.append(
                    f"Section {section_index} content page title is duplicated: {seen_page_titles[duplicate_key]}."
                )
            else:
                seen_page_titles[duplicate_key] = normalized_title
    return ValidationResult(messages=messages)
