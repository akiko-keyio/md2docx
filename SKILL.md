---
name: md2docx
description: "Convert Academic Markdown to polished Word .docx using pure Python, with native OMML equations, auto-sizing aspect-locked images, centered/justified captions, 3-line tables with vertically-centered cells, smart dashes, and bold blue cross-reference fields."
---

# md2docx (Markdown to Academic Word Converter)

This skill converts a Markdown document to a highly polished, publisher-ready academic Word document (`.docx`) without requiring external heavy tools like Pandoc.

## 1. System & Python Dependencies

This converter runs on pure Python but relies on Microsoft Office's native XSLT stylesheet (`MML2OMML.XSL`) to translate MathML into Microsoft Office's native math format (OMML).

### 1.1 Requirements
- **Microsoft Office** (specifically `MML2OMML.XSL` located in Program Files)
- Python packages (install using `requirements.txt` in this skill folder):
  ```bash
  pip install -r "<SKILL_DIR>/requirements.txt"
  ```

## 2. Command Usage

```bash
python -m scripts.md2docx <input.md> -o <output.docx> [-p academic-manuscript]
```
- `<input.md>`: Path to the manuscript.
- `-o <output.docx>`: Destination Word document.
- `-p`: Style profile (default: `academic-manuscript`).

---

## 3. Formatting & Conversion Rules (The "Secret Sauce")

When executing or modifying this converter, adhere strictly to these formatting invariants:

### 3.1 Mathematical Formula Rendering
- **Conversion Chain**: LaTeX → MathML (via `latex2mathml`) → OMML (via `MML2OMML.XSL`) → Post-Processed OMML.
- **Font & Operators**: No explicit fonts are written into math runs (letting Word use default Cambria Math). Known abbreviations/operators (e.g., `ZTD`, `RMS`, `RMSE`, `STD`, `sin`, `cos`, `exp`) are automatically post-processed to be upright (`sty val="p"`), preventing them from rendering in italic.

### 3.2 Dynamic Cross-References
- **In-text citations** matching `Eq. (N)`, `Fig. N`, and `Table N` are dynamically split and wrapped into MS Word **REF fields** pointing to bookmarks (`_EqN`, `_FigN`, `_TabN`).
- **Visual Styles**:
  - The text prefix (e.g., `Fig. ` or `Table `) is styled **bold**.
  - The reference number itself is styled **bold + pure blue (`0000FF`)**, and **not underlined** (re-written directly inside the REF field runs).
- **Captions Bookmark Wrapper**: Paragraph bookmarks wrapping the whole caption are generated as reference targets. Cross-reference fields inside captions are **explicitly disabled** (so captions don't reference themselves and turn blue).

### 3.3 Image Handling & Column-Width Hints
- **Explicit Sizing**: Controlled directly via alt-text in Markdown:
  - `![half](...)` scales the image to **single-column width (60% of content width)**.
  - `![full](...)` scales the image to **full-width (16.51 cm / 6.5 inches)**.
- **Implicit Sizing**: If alt-text is not "half" or "full", images with pixel widths ≤ 1500px are treated as `half`, others as `full`.
- **Aspect Ratio Locking**: Explicitly sets `<a:picLocks noChangeAspect="1"/>` with correct XML namespaces so users cannot distort images.

### 3.4 Table Styles & Padding
- **Three-line Table**: Drawn with top/bottom borders (1.5pt) and a header bottom border (0.75pt). Left, right, and vertical borders are removed.
- **Vertical Centering**: Every table cell is vertically centered using `<w:vAlign w:val="center"/>`.
- **Padding & Line Spacing**: Paragraphs inside table cells have no indent, are horizontally centered, and are padded with **0.25-line before and after** spacing (`beforeLines="25"`, `afterLines="25"`, `line="240"`, `lineRule="auto"`).

### 3.5 Headings Spacing
- **Level 1 & 2 Headings** (`<h1>`, `<h2>` / `#`, `##`) are configured with:
  - **Spacing Before**: 1.0 line (`beforeLines="100"`)
  - **Spacing After**: 0 pt (`after="0"`)
  - **Line Spacing**: 240 (single spacing)

### 3.6 Caption Alignment
- Caption paragraphs (both tables and figures) are centered or justified based on length:
  - **Single-line Captions** (total length ≤ 80 characters) → **Centered** (`jc="center"`).
  - **Multi-line Captions** (total length > 80 characters) → **Justified** (`jc="both"`).

### 3.7 Typographical Dash Conversions
- Typographer is turned on. Standard hyphens are converted as follows during conversion:
  - `--` → `–` (en dash, e.g., range values: `2020–2021`, `ERA5–GNSS`)
  - `---` → `—` (em dash, e.g., insertion phrases)

---

## 4. Integration with other Skills

- **Office Repack/Validation Skill (`docx` / `accept_changes`)**: After converting, you can optionally unpack or inspect the XML structures using the `office/unpack.py` script.
