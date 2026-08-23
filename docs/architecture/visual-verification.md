# Dual Verification: Structural & Visual Layout Verification

## 1. Dual Verification Architecture
A document modification is only considered successful when it passes **both** verification layers:

1. **Structural Verification**:
   - Independent reopen of the physical `.docx` archive.
   - XML schema validation and absence of corrupted tags.
   - Font family, size, line spacing, and alignment consistency.
2. **Visual Layout Verification (`VisualLayoutVerifier`)**:
   - Detection of printable margin overflows (e.g. images or diagrams wider than printable page width).
   - Table column overflow checks (>10 columns on portrait page).
   - Heading hierarchy sequence checks (detecting skipped heading levels such as H1 -> H3).
   - Detection of trailing empty blocks creating unwanted blank pages.
   - Excessive vertical spacing (>72pt).
