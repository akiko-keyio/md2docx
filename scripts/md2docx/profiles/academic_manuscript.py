"""Academic manuscript style profile.

Matches docx_format_spec.md — Times New Roman, 12pt body, XITS Math, etc.
Template: paper/manuscript/reference_doc_unpacked/
"""

PROFILE = {
    "name": "academic-manuscript",
    "description": "Academic manuscript (Times New Roman 12pt, XITS Math)",

    # Path to unpacked reference doc (provides styles.xml, theme, etc.)
    "template_dir": "paper/manuscript/reference_doc_unpacked",

    # Style ID mapping: logical name → Word styleId
    "styles": {
        "body_text":        "BodyText",
        "heading1":         "10",       # Word built-in Heading 1
        "heading2":         "20",       # Word built-in Heading 2
        "figure_block":     "FigureBlock",
        "fig_caption":      "FigCaption",
        "equation":         "Equation",
        "equation_table":   "EquationTable",
        "three_line_table": "ThreeLineTable",
    },

    # Equation table layout
    "equation": {
        "col_widths": [8706, 654],   # twips: equation col, number col
        "spacing_before": 50,         # 0.5 lines
        "spacing_after": 50,
        "number_font": "Times New Roman",
        "math_font": "Cambria Math",
    },

    # Caption rules
    "caption": {
        "short_threshold": 60,       # chars: ≤ this → center, > this → justify
    },

    # Table caption: spacing is opposite of fig caption
    "table_caption": {
        "spacing_before_lines": 50,  # 0.5 line before
        "spacing_after_lines": 0,    # 0 after
    },

    # Table cells
    "table_cell": {
        "font_size": 21,             # 10.5pt (half-points)
        "first_line_indent": 0,
    },

    # Page layout (US Letter per original manuscript)
    "page": {
        "width": 12240,              # twips (8.5 inches)
        "height": 15840,             # twips (11 inches)
        "margin_top": 1440,          # 1 inch
        "margin_bottom": 1440,
        "margin_left": 1440,
        "margin_right": 1440,
        "content_width_cm": 16.51,   # 6.5 inches = 16.51 cm (Letter 8.5" - 2×1" margins)
        "content_width_twips": 9360, # 12240 - 1440*2
        "line_numbering": True,      # continuous line numbers
    },
}
