# 08 - Align First Version Format And Output Verification

**What to build:** Improve the first generated PPTX so it more closely matches the reference training document structure, then add verification that catches obvious output problems before users rely on the file.

**Blocked by:** 07 - Connect UI To PPTX Generation.

**Status:** ready-for-agent

- [ ] Generated slides use a 16:9 slide size matching the reference deck.
- [ ] Cover, table of contents, section, and content pages use consistent placement and hierarchy.
- [ ] Generated typography and spacing are close enough to the reference deck for a first company-internal draft.
- [ ] Verification confirms the generated file can be opened as a PPTX.
- [ ] Verification confirms the expected slide count.
- [ ] Verification confirms required text exists on the expected slides.
- [ ] Verification confirms expected navigation links exist.
- [ ] Verification failure is reported clearly to the user or developer.
