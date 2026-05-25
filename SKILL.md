---
name: md2docx
description: "Convert Markdown to Word .docx with configurable style profiles (e.g., academic manuscript)."
---

# md2docx

Pure-Python Markdown → Word converter.

## 1. Requirements

- **Microsoft Office** — provides `MML2OMML.XSL` for equation rendering
- **Python packages**: `pip install -r "<SKILL_DIR>/requirements.txt"`

## 2. Usage

```bash
python -m md2docx <input.md> -o <output.docx> [--style academic-manuscript]
```

Run from `<SKILL_DIR>`. The skill is self-contained.

## 3. Architecture

```
<SKILL_DIR>/
├── md2docx/              ← Python package
│   ├── cli.py            ← argument parsing
│   ├── convert.py        ← orchestrator: parse → build → package
│   ├── parse.py          ← Markdown → block list
│   ├── build.py          ← block list → document.xml
│   └── math.py           ← LaTeX → OMML
└── styles/               ← style presets (each = unpacked .docx)
    └── academic-manuscript/
        └── word/
            ├── styles.xml      ← visual formatting
            ├── settings.xml    ← document settings
            └── theme/          ← color/font theme
```

**Flow**: `convert.py` copies `styles/<name>/` as output skeleton → `parse.py` reads Markdown → `build.py` generates `document.xml` → zip as `.docx`.

## 4. Input Contract

Write Markdown following these patterns to trigger correct conversion:

| Pattern | Behavior |
|---------|----------|
| `$$...\tag{N}$$` | Numbered equation in 2-column table |
| `![half](path)` or image ≤1500px | 60% width |
| `![full](path)` or image >1500px | Full text width |
| `**Fig. N** caption` | Figure caption with bookmark |
| `**Table N.** caption` | Table caption with bookmark |
| `Eq. (N)`, `Fig. N`, `Table N` in text | Cross-reference REF field |
| GFM pipe table | Three-line table |
| `--` / `---` | En-dash / em-dash |

## 5. Styles

Each style = a directory under `styles/` containing Word Open XML (the internal structure of a .docx file). The `word/` subdirectory is mandated by the Open XML specification.

- **Visual formatting** (fonts, colors, spacing) → edit `word/styles.xml`
- **Page layout** (dimensions, margins) → written into `sectPr` by `build.py`

To create a new style: copy `styles/academic-manuscript/` → rename → edit `word/styles.xml`.

## 6. Troubleshooting

| Symptom | Fix |
|---------|-----|
| `PermissionError` on output | Close the file in Word |
| Equation shows literal `\arg` | Use `\operatorname{arg}` |
| Image too small | Check pixel width (≤1500px → half) |
| DOCX opens with repair prompt | Stale comment XML (auto-cleaned) |
