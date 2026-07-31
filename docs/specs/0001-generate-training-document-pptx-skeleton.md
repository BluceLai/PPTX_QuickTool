# Spec: Generate Training Document PPTX Skeleton

## Problem Statement

Company users need to create PowerPoint-based training documents that follow a shared internal structure. Today, creating these documents requires repeated manual work: copying slides, adding section pages, updating the table of contents, setting navigation links, and keeping the page order consistent.

This slows down first-draft creation and makes formatting inconsistent across documents. The first version of PPTX QuickTool should reduce this setup work by generating a complete PPTX skeleton from structured user input.

## Solution

Build a simple desktop tool with a UI. The user enters the training document title, table of contents items, section titles, and optional content page titles. The tool generates a new PPTX that follows the company's training document structure.

The generated deck should include a title slide, a table of contents slide, section start slides, and content pages. It should create navigation links from the table of contents to each section and from each section start page back to the table of contents where the template supports it. Content pages should stay clean and should not display a return-to-contents link.

The goal is to quickly produce a clean first draft, not to complete the final training content.

## User Stories

1. As a company user, I want to create a new training document from a simple form, so that I do not need to manually copy an existing PPTX.
2. As a company user, I want to enter the PPT title, so that the generated deck has the correct cover slide.
3. As a company user, I want to enter table of contents items, so that the generated deck has a complete agenda slide.
4. As a company user, I want each table of contents item to become a section in the deck, so that the document structure matches the planned training flow.
5. As a company user, I want to enter optional content page titles under each section, so that the tool can create the correct number of content pages when they are needed.
6. As a company user, I want the tool to generate section start pages, so that each chapter begins consistently.
7. As a company user, I want the tool to generate blank or lightly prepared content pages, so that I can fill in details after the first draft is created.
8. As a company user, I want the table of contents entries to link to their matching section pages, so that readers can quickly jump to the right chapter.
9. As a company user, I want section start pages to have a link back to the table of contents, so that readers can navigate the file quickly.
10. As a company user, I want generated pages to follow the reference training document format, so that files remain consistent across the company.
11. As a company user, I want the output to be a new PPTX file, so that the source template or reference file is not changed.
12. As a company user, I want to preview or review the planned page list before generation, so that I can catch missing sections before creating the PPTX.
13. As a company user, I want to reorder sections before generation, so that the output follows the training sequence I intend.
14. As a company user, I want to add, rename, and remove sections in the UI, so that I can adjust the structure without editing raw data.
15. As a company user, I want to add, rename, and remove content pages within a section, so that the generated deck matches the planned document.
16. As a company user, I want validation for missing titles or empty sections, so that the generated deck is not confusing.
17. As a company user, I want clear generation success and error messages, so that I know whether the output file is ready.
18. As a company user, I want to choose where the generated PPTX is saved, so that I can place it in the right project folder.
19. As a company user, I want the tool to preserve the expected slide size and layout style, so that the output feels like the reference document.
20. As a company user, I want to choose a PPTX template source when needed, so that generated files can follow the right department format without changing the template file.
21. As a company user, I want the app to start with a default foreword section, so that a common training document chapter is ready to edit.
22. As a company user, I want long table-of-contents slides to wrap into up to three columns, so that agenda text does not run past the slide boundary.
23. As a future maintainer, I want the PPTX generation logic separated from the desktop UI, so that it can be tested without clicking through the app.

## Implementation Decisions

- The first version will focus on generating the training document structure, not detailed teaching content.
- The product will be a desktop tool with a simple graphical UI.
- The user input model will represent a document title, an ordered list of sections, and an optional ordered list of content pages for each section.
- The generated PPTX will include title, table of contents, section start, and content page slide types.
- The table of contents should be generated from the section list and each section's content page titles rather than manually typed independently.
- Long table-of-contents and section agenda content should wrap into one, two, or three columns before exceeding the slide body area.
- Hyperlinks are part of the first version because saving link setup time is a core goal.
- Table of contents section entries link to their matching section start pages. Section start pages link back to the table of contents. Content pages do not display a return-to-contents link.
- The source/reference PPTX should be treated as a formatting reference or template source, not as an in-place file to mutate.
- The app includes a default training document template and can optionally generate from a user-selected PPTX template source.
- The generated PPTX should be exported as a new file.
- Screenshot-heavy teaching pages, mouse markers, arrows, and screenshot annotation workflows are not part of the first version.
- The core test seam is the document structure generator: user input becomes a deterministic page plan.
- The second test seam is the navigation planner: page plan becomes deterministic table-of-contents and return-link relationships.
- The third test seam is the PPTX exporter: page plan plus navigation relationships become a valid PPTX file.

## Testing Decisions

- Tests should focus on external behavior: given user input, the system produces the expected page plan, link plan, and PPTX output.
- Unit tests should verify that a title plus section/page input creates the correct page order.
- Unit tests should verify that table of contents entries point to the correct generated section pages.
- Unit tests should verify that generated section start pages can link back to the table of contents and generated content pages do not display return links.
- Validation tests should cover missing document titles, empty section titles, duplicate or blank content page titles, and the allowed case where a section has no content pages.
- UI tests should verify the default foreword section.
- PPTX generation tests should verify that long agenda content uses no more than three columns and stays within the slide boundary.
- PPTX generation tests should verify that the output file opens as a valid PPTX and contains the expected slide count and visible text.
- UI tests can be added after the first UI framework is selected; the first priority is testing the underlying generation behavior.
- Visual fidelity tests should compare generated slide structure against the reference format at a basic level, such as slide size, expected layouts, and required text placement.

## Out of Scope

- Creating detailed screenshot-based teaching pages.
- Annotating screenshots with arrows, mouse icons, or numbered callouts.
- Automatically writing the full body content of the training document.
- Translating training content.
- Importing existing PPTX files and restructuring them automatically.
- Managing a template catalog or automatically selecting among multiple company templates.
- Publishing directly to SharePoint, Teams, or other document systems.
- Collaborative editing.

## Further Notes

The reference PPTX shows a training-document pattern with a title slide, table of contents, section-oriented pages, and content pages. The first product slice should reproduce the structural convenience of that pattern: page creation, ordering, and navigation.

The most important product promise is speed: users should be able to enter the planned document structure and get a usable PPTX first draft in minutes.
