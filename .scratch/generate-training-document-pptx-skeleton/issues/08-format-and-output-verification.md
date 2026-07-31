# 08 - Align First Version Format And Output Verification

**What to build:** Improve the first generated PPTX so it more closely matches the reference training document structure, then add verification that catches obvious output problems before users rely on the file.

**Blocked by:** 07 - Connect UI To PPTX Generation.

**Status:** done

- [x] Generated slides use a 16:9 slide size matching the reference deck.
- [x] Cover, table of contents, section, and content pages use consistent placement and hierarchy.
- [x] Generated typography and spacing are close enough to the reference deck for a first company-internal draft.
- [x] Verification confirms the generated file can be opened as a PPTX.
- [x] Verification confirms the expected slide count.
- [x] Verification confirms required text exists on the expected slides.
- [x] Verification confirms expected navigation links exist.
- [x] Verification failure is reported clearly to the user or developer.
