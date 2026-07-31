# 03 - Generate PPTX Page Plan

**What to build:** Convert valid training document input into a deterministic page plan. The plan should describe the exact generated slide order before any PPTX file is created.

**Blocked by:** 02 - Define Training Document Input Model.

**Status:** ready-for-agent

- [ ] A valid input produces one cover page in the page plan.
- [ ] A valid input produces one table of contents page in the page plan.
- [ ] Each section produces one section start page.
- [ ] Each content page title produces one content page under its section.
- [ ] Page order is deterministic and matches the entered section order.
- [ ] The page plan includes enough stable identifiers to support navigation links later.
- [ ] Tests cover multiple sections with different numbers of content pages.
