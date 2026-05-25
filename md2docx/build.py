"""Word XML document builder — converts structured blocks to document.xml.

Reads a style profile and generates spec-compliant Word Open XML.
"""

import re
from pathlib import Path
from lxml import etree
from PIL import Image

from .math import MathEngine

# ── Namespaces ────────────────────────────────────────────────────────
NS = {
    'w':    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm':    'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'r':    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':   'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':    'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':  'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'w14':  'http://schemas.microsoft.com/office/word/2010/wordml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'mc':   'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wpc':  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'o':    'urn:schemas-microsoft-com:office:office',
    'v':    'urn:schemas-microsoft-com:vml',
    'w10':  'urn:schemas-microsoft-com:office:word',
}

for _pfx, _uri in NS.items():
    etree.register_namespace(_pfx, _uri)

W  = NS['w']
M  = NS['m']
R  = NS['r']
WP = NS['wp']
A  = NS['a']
PIC = NS['pic']

# EMU per cm
EMU_PER_CM = 360000


def _qn(tag: str) -> str:
    pfx, local = tag.split(':')
    return f'{{{NS[pfx]}}}{local}'


class DocxBuilder:
    """Build a Word document.xml from parsed blocks + style profile."""

    def __init__(self, profile: dict, math_engine: MathEngine):
        self.profile = profile
        self.math = math_engine
        self._images: list[dict] = []     # {"path": ..., "rId": ..., "media_name": ...}
        self._rid_counter = 10            # start relationship IDs at rId10
        self._doc_id_counter = 1          # unique IDs for drawings
        self._bookmark_id = 100           # bookmark IDs for cross-refs

    # ── Public API ────────────────────────────────────────────────────

    def build(self, blocks: list[dict]) -> etree._Element:
        """Build complete <w:document> element from blocks."""
        doc = etree.Element(_qn('w:document'))
        doc.set(f'xmlns:{k}' if k != 'w' else 'xmlns:w',  # handled by register
                v) if False else None  # no-op; namespaces registered globally

        body = etree.SubElement(doc, _qn('w:body'))

        for block in blocks:
            btype = block["type"]
            if btype == "heading":
                self._add_heading(body, block)
            elif btype == "paragraph":
                self._add_paragraph(body, block)
            elif btype == "display_math":
                self._add_equation_table(body, block)
            elif btype == "figure":
                self._add_figure(body, block)
            elif btype == "caption":
                self._add_caption(body, block)
            elif btype == "table":
                self._add_data_table(body, block)

        # Section properties (page layout)
        self._add_section_props(body)
        return doc

    @property
    def images(self) -> list[dict]:
        """Return list of images that need to be embedded."""
        return self._images

    # ── Helpers ───────────────────────────────────────────────────────

    def _next_rid(self) -> str:
        rid = f'rId{self._rid_counter}'
        self._rid_counter += 1
        return rid

    def _next_doc_id(self) -> int:
        did = self._doc_id_counter
        self._doc_id_counter += 1
        return did

    def _style(self, key: str) -> str:
        """Look up a styleId from the profile."""
        return self.profile["styles"][key]

    # ── Paragraph construction ────────────────────────────────────────

    def _make_p(self, parent, style_key: str, runs=None, jc=None, extra_pPr=None, refs_enabled=True) -> etree._Element:
        """Create a <w:p> with style and optional runs."""
        p = etree.SubElement(parent, _qn('w:p'))
        pPr = etree.SubElement(p, _qn('w:pPr'))
        etree.SubElement(pPr, _qn('w:pStyle')).set(_qn('w:val'), self._style(style_key))
        if extra_pPr:
            extra_pPr(pPr)          # spacing/indent must come before jc in OOXML
        if jc:
            etree.SubElement(pPr, _qn('w:jc')).set(_qn('w:val'), jc)
        if runs:
            self._add_runs(p, runs, refs_enabled=refs_enabled)
        return p

    # Regexes for cross-references in body text
    _EQ_REF_RE    = re.compile(r'Eq\.\s*\((\d+)\)')        # Eq. (N)
    _EQ_REF_SPLIT = re.compile(r'Eq\.\s*\(\d+\)')          # no-capture split
    _FIG_REF_RE   = re.compile(r'(Fig|Table)\.\s*(\d+)')     # Fig. N / Table N
    _FIG_REF_SPLIT= re.compile(r'(?:Fig|Table)\.\s*\d+')    # no-capture split

    def _add_runs(self, parent, runs: list[dict], refs_enabled=True):
        """Append run elements to a paragraph."""
        for run in runs:
            rtype = run.get("type", "text")
            if rtype == "text":
                if refs_enabled:
                    self._add_text_run_with_refs(parent, run)
                else:
                    self._add_text_run(parent, run)
            elif rtype == "inline_math":
                self._add_inline_math(parent, run)
            elif rtype == "break":
                r = etree.SubElement(parent, _qn('w:r'))
                etree.SubElement(r, _qn('w:br'))

    def _add_text_run_with_refs(self, parent, run: dict):
        """Add text run, replacing 'Eq. (N)', 'Fig. N', 'Table N' with REF fields."""
        text = run.get("text", "")
        if not text:
            return
        bold = run.get("bold", False)
        italic = run.get("italic", False)

        # Split on ALL cross-reference patterns together
        _ALL_SPLIT = re.compile(r'Eq\.\s*\(\d+\)|(?:Fig|Table)\.\s*\d+')
        parts = _ALL_SPLIT.split(text)
        # Collect matches in order: (kind, number)
        matches = []
        for m in re.finditer(r'(Eq\.\s*\((\d+)\))|((Fig|Table)\.\s*(\d+))', text):
            if m.group(1):   # Eq. (N)
                matches.append(('eq', m.group(2)))
            else:            # Fig. N or Table N
                matches.append((m.group(4).lower(), m.group(5)))

        for i, part in enumerate(parts):
            if part:
                self._add_text_run(parent, {"type": "text", "text": part,
                                            "bold": bold, "italic": italic})
            if i < len(matches):
                kind, num = matches[i]
                if kind == 'eq':
                    self._add_eq_ref_field(parent, num, bold, italic)
                else:
                    bm_name = f'_Fig{num}' if kind == 'fig' else f'_Tab{num}'
                    label   = f'Fig. {num}' if kind == 'fig' else f'Table {num}'
                    # Force bold; number gets blue color inside
                    self._add_generic_ref_field(parent, bm_name, label, bold=True, italic=italic)

    def _add_generic_ref_field(self, parent, bm_name: str, display: str,
                                bold=False, italic=False):
        """Insert a REF field: prefix text bold, number in hyperlink-blue.

        display is e.g. 'Fig. 3' or 'Table 1'.
        The prefix ('Fig. ' / 'Table ') is bold-only; the number is bold + blue.
        """
        # Split display into prefix + number, e.g. 'Fig. ' and '3'
        m = re.match(r'^(.*?)(\d+)$', display)
        prefix = m.group(1) if m else display
        number = m.group(2) if m else ''

        # --- Prefix run (bold, normal colour) ---
        if prefix:
            r_pre = etree.SubElement(parent, _qn('w:r'))
            rPr_pre = etree.SubElement(r_pre, _qn('w:rPr'))
            etree.SubElement(rPr_pre, _qn('w:b'))
            if italic: etree.SubElement(rPr_pre, _qn('w:i'))
            t_pre = etree.SubElement(r_pre, _qn('w:t'))
            t_pre.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t_pre.text = prefix

        # --- Field begin ---
        r1 = etree.SubElement(parent, _qn('w:r'))
        fldChar1 = etree.SubElement(r1, _qn('w:fldChar'))
        fldChar1.set(_qn('w:fldCharType'), 'begin')

        r2 = etree.SubElement(parent, _qn('w:r'))
        instrText = etree.SubElement(r2, _qn('w:instrText'))
        instrText.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        instrText.text = f' REF {bm_name} \\h '

        r3 = etree.SubElement(parent, _qn('w:r'))
        fldChar2 = etree.SubElement(r3, _qn('w:fldChar'))
        fldChar2.set(_qn('w:fldCharType'), 'separate')

        # --- Display run: bold + explicit hyperlink-blue, no underline ---
        r4 = etree.SubElement(parent, _qn('w:r'))
        rPr4 = etree.SubElement(r4, _qn('w:rPr'))
        etree.SubElement(rPr4, _qn('w:b'))
        if italic: etree.SubElement(rPr4, _qn('w:i'))
        _color = etree.SubElement(rPr4, _qn('w:color'))
        _color.set(_qn('w:val'), '0000FF')  # classic hyperlink blue
        t4 = etree.SubElement(r4, _qn('w:t'))
        t4.text = number

        r5 = etree.SubElement(parent, _qn('w:r'))
        fldChar3 = etree.SubElement(r5, _qn('w:fldChar'))
        fldChar3.set(_qn('w:fldCharType'), 'end')

    def _add_eq_ref_field(self, parent, eq_num: str, bold=False, italic=False):
        """Insert a REF field for equation cross-reference: Eq. (N)."""
        bm_name = f'_Eq{eq_num}'

        # "Eq. " prefix as normal run
        self._add_text_run(parent, {"type": "text", "text": "Eq. ",
                                    "bold": bold, "italic": italic})

        # Field begin
        r1 = etree.SubElement(parent, _qn('w:r'))
        fldChar1 = etree.SubElement(r1, _qn('w:fldChar'))
        fldChar1.set(_qn('w:fldCharType'), 'begin')

        # Field instruction
        r2 = etree.SubElement(parent, _qn('w:r'))
        instrText = etree.SubElement(r2, _qn('w:instrText'))
        instrText.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        instrText.text = f' REF {bm_name} \\h '

        # Field separate
        r3 = etree.SubElement(parent, _qn('w:r'))
        fldChar2 = etree.SubElement(r3, _qn('w:fldChar'))
        fldChar2.set(_qn('w:fldCharType'), 'separate')

        # Field display text
        r4 = etree.SubElement(parent, _qn('w:r'))
        if bold or italic:
            rPr = etree.SubElement(r4, _qn('w:rPr'))
            if bold:
                etree.SubElement(rPr, _qn('w:b'))
            if italic:
                etree.SubElement(rPr, _qn('w:i'))
        t4 = etree.SubElement(r4, _qn('w:t'))
        t4.text = f'({eq_num})'

        # Field end
        r5 = etree.SubElement(parent, _qn('w:r'))
        fldChar3 = etree.SubElement(r5, _qn('w:fldChar'))
        fldChar3.set(_qn('w:fldCharType'), 'end')

    def _add_text_run(self, parent, run: dict):
        text = run.get("text", "")
        if not text:
            return
        r = etree.SubElement(parent, _qn('w:r'))
        rPr = etree.SubElement(r, _qn('w:rPr'))
        if run.get("bold"):
            etree.SubElement(rPr, _qn('w:b'))
            etree.SubElement(rPr, _qn('w:bCs'))
        if run.get("italic"):
            etree.SubElement(rPr, _qn('w:i'))
            etree.SubElement(rPr, _qn('w:iCs'))
        t = etree.SubElement(r, _qn('w:t'))
        t.text = text
        if text.startswith(' ') or text.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    def _add_inline_math(self, parent, run: dict):
        """Insert inline OMML math."""
        try:
            omath = self.math.to_omml(run["latex"])
            # Ensure we have an m:oMath element
            if omath.tag == f'{{{M}}}oMathPara':
                omath_inner = omath.find(f'{{{M}}}oMath')
                if omath_inner is not None:
                    omath = omath_inner
            parent.append(omath)
        except Exception:
            # Fallback: render as plain text
            r = etree.SubElement(parent, _qn('w:r'))
            t = etree.SubElement(r, _qn('w:t'))
            t.text = f'${run["latex"]}$'

    # ── Block builders ────────────────────────────────────────────────

    def _add_heading(self, body, block: dict):
        level = block["level"]
        key = f'heading{level}'
        if level in (1, 2):
            def _heading_spacing(pPr):
                sp = etree.SubElement(pPr, _qn('w:spacing'))
                sp.set(_qn('w:beforeLines'), '100')  # 1 line before
                sp.set(_qn('w:after'), '0')           # 0 pt after
                sp.set(_qn('w:line'), '240')
                sp.set(_qn('w:lineRule'), 'auto')
            self._make_p(body, key, block["runs"], extra_pPr=_heading_spacing)
        else:
            self._make_p(body, key, block["runs"])

    def _add_paragraph(self, body, block: dict):
        self._make_p(body, "body_text", block["runs"])

    def _add_caption(self, body, block: dict):
        """Add a figure/table caption paragraph.
        
        Bold rules:
        - Fig captions: only 'Fig. N' is bold, description is normal weight
        - Table captions: all runs bold
        Short captions → center; long captions → justify.
        """
        ref_type = block.get("ref_type", "fig")

        if ref_type == "table":
            # Table: all bold
            runs = []
            for r in block["runs"]:
                if r.get("type") == "text":
                    runs.append({**r, "bold": True})
                else:
                    runs.append(r)
        else:
            # Fig: only the first non-empty bold run stays bold, rest normal
            runs = []
            label_done = False
            for r in block["runs"]:
                if r.get("type") != "text":
                    runs.append(r)
                    continue
                if not label_done and r.get("bold"):
                    runs.append({**r, "bold": True})
                    label_done = True
                else:
                    runs.append({**r, "bold": False})

        # Short caption (≈ single line) → center; long (multi-line) → justify
        total_len = sum(len(r.get("text", "")) for r in runs if r.get("type") == "text")
        threshold = self.profile.get("caption", {}).get("short_threshold", 80)
        jc = "center" if total_len <= threshold else "both"

        # Spacing differs: table caption has before=0.5, after=0; fig has before=0, after=0.5
        def extra_pPr(pPr):
            sp = etree.SubElement(pPr, _qn('w:spacing'))
            if ref_type == "table":
                tbl_cap = self.profile.get("table_caption", {})
                sp.set(_qn('w:beforeLines'), str(tbl_cap.get("spacing_before_lines", 50)))
                sp.set(_qn('w:afterLines'), str(tbl_cap.get("spacing_after_lines", 0)))
            else:
                sp.set(_qn('w:before'), '0')
                sp.set(_qn('w:afterLines'), '50')
            sp.set(_qn('w:line'), '300')
            sp.set(_qn('w:lineRule'), 'auto')

        # Build the caption paragraph — no cross-reference REF fields inside captions
        p = self._make_p(body, "fig_caption", runs, jc=jc, extra_pPr=extra_pPr, refs_enabled=False)

        # Add bookmark wrapping ONLY the number digit in the caption.
        # e.g. for "Fig. 3 caption…", bookmark wraps just "3" so that
        # REF _Fig3 \h resolves to "3", not "Fig. 3 caption…".
        cap_text = ''.join(r.get('text', '') for r in block['runs'] if r.get('type') == 'text')
        m = re.match(r'(Fig|Table)\.?\s*(\d+)', cap_text)
        if m:
            import copy
            kind, num = m.group(1), m.group(2)
            bm_name = f'_Fig{num}' if kind == 'Fig' else f'_Tab{num}'
            bm_id   = str(self._bookmark_id)
            self._bookmark_id += 1

            bs = etree.Element(_qn('w:bookmarkStart'))
            bs.set(_qn('w:id'), bm_id)
            bs.set(_qn('w:name'), bm_name)
            be = etree.Element(_qn('w:bookmarkEnd'))
            be.set(_qn('w:id'), bm_id)

            # Find the first run whose text contains the label (e.g. "Fig. 3")
            # and split it into [prefix_run][bookmarkStart][number_run][bookmarkEnd]
            inserted = False
            for i, child in enumerate(list(p)):
                if child.tag != _qn('w:r'):
                    continue
                t_el = child.find(_qn('w:t'))
                if t_el is None or not t_el.text:
                    continue
                nm = re.match(r'^((?:Fig|Table)\.?\s*)(\d+)(.*)', t_el.text)
                if nm:
                    prefix_txt = nm.group(1)   # "Fig. "
                    number_txt = nm.group(2)   # "3"
                    suffix_txt = nm.group(3)   # ""

                    # Shorten existing run to just the prefix
                    t_el.text = prefix_txt
                    if prefix_txt.endswith(' ') or prefix_txt.startswith(' '):
                        t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

                    # Build separate number run (same rPr)
                    r_num = etree.Element(_qn('w:r'))
                    rPr_orig = child.find(_qn('w:rPr'))
                    if rPr_orig is not None:
                        r_num.append(copy.deepcopy(rPr_orig))
                    t_num = etree.SubElement(r_num, _qn('w:t'))
                    t_num.text = number_txt + suffix_txt

                    # Insert after existing prefix run: [bs][r_num][be]
                    p.insert(i + 1, bs)
                    p.insert(i + 2, r_num)
                    p.insert(i + 3, be)
                    inserted = True
                    break

            if not inserted:
                # Fallback: wrap whole paragraph
                pPr_idx = list(p).index(p.find(_qn('w:pPr')))
                p.insert(pPr_idx + 1, bs)
                p.append(be)

    # ── Equation table ────────────────────────────────────────────────

    def _add_equation_table(self, body, block: dict):
        """Wrap display math in a 2-column EquationTable."""
        eq_cfg = self.profile.get("equation", {})
        col_widths = eq_cfg.get("col_widths", [8706, 654])

        tbl = etree.SubElement(body, _qn('w:tbl'))

        # tblPr
        tblPr = etree.SubElement(tbl, _qn('w:tblPr'))
        etree.SubElement(tblPr, _qn('w:tblStyle')).set(
            _qn('w:val'), self._style("equation_table"))
        tblW = etree.SubElement(tblPr, _qn('w:tblW'))
        tblW.set(_qn('w:w'), '0')
        tblW.set(_qn('w:type'), 'auto')

        # Borders all none
        tblBorders = etree.SubElement(tblPr, _qn('w:tblBorders'))
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = etree.SubElement(tblBorders, f'{{{W}}}{side}')
            b.set(_qn('w:val'), 'none')
            b.set(_qn('w:sz'), '0')
            b.set(_qn('w:space'), '0')
            b.set(_qn('w:color'), 'auto')

        etree.SubElement(tblPr, _qn('w:tblLayout')).set(_qn('w:type'), 'fixed')
        tblLook = etree.SubElement(tblPr, _qn('w:tblLook'))
        tblLook.set(_qn('w:val'), '04A0')
        for attr, val in [('firstRow','1'),('lastRow','0'),('firstColumn','1'),
                          ('lastColumn','0'),('noHBand','0'),('noVBand','1')]:
            tblLook.set(_qn(f'w:{attr}'), val)

        # Grid
        tblGrid = etree.SubElement(tbl, _qn('w:tblGrid'))
        for cw in col_widths:
            etree.SubElement(tblGrid, _qn('w:gridCol')).set(_qn('w:w'), str(cw))

        # Row
        tr = etree.SubElement(tbl, _qn('w:tr'))

        # Left cell — equation
        tc1 = etree.SubElement(tr, _qn('w:tc'))
        tcPr1 = etree.SubElement(tc1, _qn('w:tcPr'))
        tcW1 = etree.SubElement(tcPr1, _qn('w:tcW'))
        tcW1.set(_qn('w:w'), str(col_widths[0]))
        tcW1.set(_qn('w:type'), 'dxa')
        etree.SubElement(tcPr1, _qn('w:vAlign')).set(_qn('w:val'), 'center')

        p1 = etree.SubElement(tc1, _qn('w:p'))
        pPr1 = etree.SubElement(p1, _qn('w:pPr'))
        etree.SubElement(pPr1, _qn('w:pStyle')).set(
            _qn('w:val'), self._style("equation"))
        sp1 = etree.SubElement(pPr1, _qn('w:spacing'))
        sp1.set(_qn('w:beforeLines'), str(eq_cfg.get("spacing_before", 50)))
        sp1.set(_qn('w:afterLines'), str(eq_cfg.get("spacing_after", 50)))
        sp1.set(_qn('w:line'), '300')
        sp1.set(_qn('w:lineRule'), 'auto')

        # Convert LaTeX → OMML and append
        try:
            omml = self.math.to_omml_para(block["latex"])
            p1.append(omml)
        except Exception as e:
            # Fallback: plain text
            r = etree.SubElement(p1, _qn('w:r'))
            t = etree.SubElement(r, _qn('w:t'))
            t.text = block["latex"]

        # Right cell — equation number
        tc2 = etree.SubElement(tr, _qn('w:tc'))
        tcPr2 = etree.SubElement(tc2, _qn('w:tcPr'))
        tcW2 = etree.SubElement(tcPr2, _qn('w:tcW'))
        tcW2.set(_qn('w:w'), str(col_widths[1]))
        tcW2.set(_qn('w:type'), 'dxa')
        etree.SubElement(tcPr2, _qn('w:vAlign')).set(_qn('w:val'), 'center')

        p2 = etree.SubElement(tc2, _qn('w:p'))
        pPr2 = etree.SubElement(p2, _qn('w:pPr'))
        etree.SubElement(pPr2, _qn('w:pStyle')).set(
            _qn('w:val'), self._style("equation"))
        etree.SubElement(pPr2, _qn('w:jc')).set(_qn('w:val'), 'right')

        if block.get("tag"):
            tag_num = block["tag"]
            bm_name = f'_Eq{tag_num}'
            bm_id = str(self._bookmark_id)
            self._bookmark_id += 1

            # Bookmark start
            bs = etree.SubElement(p2, _qn('w:bookmarkStart'))
            bs.set(_qn('w:id'), bm_id)
            bs.set(_qn('w:name'), bm_name)

            r2 = etree.SubElement(p2, _qn('w:r'))
            rPr2 = etree.SubElement(r2, _qn('w:rPr'))
            rf2 = etree.SubElement(rPr2, _qn('w:rFonts'))
            rf2.set(_qn('w:ascii'), 'Times New Roman')
            rf2.set(_qn('w:hAnsi'), 'Times New Roman')
            t2 = etree.SubElement(r2, _qn('w:t'))
            t2.text = f'({tag_num})'

            # Bookmark end
            be = etree.SubElement(p2, _qn('w:bookmarkEnd'))
            be.set(_qn('w:id'), bm_id)

    # ── Figure ────────────────────────────────────────────────────────

    def _add_figure(self, body, block: dict):
        """Add a centered figure paragraph with embedded image."""
        img_path = Path(block["path"])
        if not img_path.exists():
            # Fallback: placeholder text
            self._make_p(body, "figure_block",
                         [{"type": "text", "text": f"[Image: {img_path.name}]"}],
                         jc="center")
            return

        # Get image dimensions — preserve aspect ratio, cap at page width
        width_hint = block.get("width_hint", "full")
        page_cfg = self.profile.get("page", {})
        max_w_cm = page_cfg.get("content_width_cm", 15.24)
        max_w_emu = int(max_w_cm * EMU_PER_CM)

        with Image.open(img_path) as im:
            orig_w_px, orig_h_px = im.size
            dpi = im.info.get('dpi', (150, 150))
            if isinstance(dpi, tuple):
                dpi_x = dpi[0] if dpi[0] > 0 else 150
                dpi_y = dpi[1] if dpi[1] > 0 else 150
            else:
                dpi_x = dpi_y = dpi if dpi > 0 else 150

        # Original size in EMU (based on DPI)
        orig_w_emu = int(orig_w_px / dpi_x * 914400)  # 914400 EMU per inch
        orig_h_emu = int(orig_h_px / dpi_y * 914400)

        # Determine target width
        if width_hint == "half":
            # Half-column: 60% of text width, scale proportionally
            target_w_emu = int(max_w_emu * 0.60)
        else:
            # Full-width: always fill entire text width
            target_w_emu = max_w_emu

        # Scale height proportionally to preserve aspect ratio
        scale = target_w_emu / orig_w_emu if orig_w_emu > 0 else 1.0
        target_h_emu = int(orig_h_emu * scale)

        # Register image
        rid = self._next_rid()
        media_name = f'image{len(self._images) + 1}{img_path.suffix}'
        self._images.append({
            "path": str(img_path),
            "rId": rid,
            "media_name": media_name,
        })

        # Build paragraph with drawing
        p = etree.SubElement(body, _qn('w:p'))
        pPr = etree.SubElement(p, _qn('w:pPr'))
        etree.SubElement(pPr, _qn('w:pStyle')).set(
            _qn('w:val'), self._style("figure_block"))
        etree.SubElement(pPr, _qn('w:jc')).set(_qn('w:val'), 'center')

        r = etree.SubElement(p, _qn('w:r'))
        drawing = etree.SubElement(r, _qn('w:drawing'))
        inline = etree.SubElement(drawing, _qn('wp:inline'))
        inline.set('distT', '0')
        inline.set('distB', '0')
        inline.set('distL', '0')
        inline.set('distR', '0')

        extent = etree.SubElement(inline, _qn('wp:extent'))
        extent.set('cx', str(target_w_emu))
        extent.set('cy', str(target_h_emu))

        doc_pr = etree.SubElement(inline, _qn('wp:docPr'))
        did = self._next_doc_id()
        doc_pr.set('id', str(did))
        doc_pr.set('name', f'Picture {did}')

        graphic = etree.SubElement(inline, _qn('a:graphic'))
        graphic_data = etree.SubElement(graphic, _qn('a:graphicData'))
        graphic_data.set('uri', PIC)

        pic_el = etree.SubElement(graphic_data, f'{{{PIC}}}pic')

        nvPicPr = etree.SubElement(pic_el, f'{{{PIC}}}nvPicPr')
        cNvPr = etree.SubElement(nvPicPr, f'{{{PIC}}}cNvPr')
        cNvPr.set('id', str(did))
        cNvPr.set('name', media_name)
        cNvPicPr = etree.SubElement(nvPicPr, f'{{{PIC}}}cNvPicPr')
        # Use explicit namespace URI so lxml serialises correctly
        A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        picLocks = etree.SubElement(cNvPicPr, f'{{{A_NS}}}picLocks')
        picLocks.set('noChangeAspect', '1')

        blipFill = etree.SubElement(pic_el, f'{{{PIC}}}blipFill')
        blip = etree.SubElement(blipFill, _qn('a:blip'))
        blip.set(f'{{{R}}}embed', rid)
        stretch = etree.SubElement(blipFill, _qn('a:stretch'))
        etree.SubElement(stretch, _qn('a:fillRect'))

        spPr = etree.SubElement(pic_el, f'{{{PIC}}}spPr')
        xfrm = etree.SubElement(spPr, _qn('a:xfrm'))
        off = etree.SubElement(xfrm, _qn('a:off'))
        off.set('x', '0')
        off.set('y', '0')
        ext = etree.SubElement(xfrm, _qn('a:ext'))
        ext.set('cx', str(target_w_emu))
        ext.set('cy', str(target_h_emu))
        prstGeom = etree.SubElement(spPr, _qn('a:prstGeom'))
        prstGeom.set('prst', 'rect')

    # ── Data table (three-line) ───────────────────────────────────────

    def _add_data_table(self, body, block: dict):
        """Add a three-line data table."""
        headers = block["headers"]
        rows = block["rows"]
        n_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
        if n_cols == 0:
            return

        tbl = etree.SubElement(body, _qn('w:tbl'))

        # tblPr
        tblPr = etree.SubElement(tbl, _qn('w:tblPr'))
        etree.SubElement(tblPr, _qn('w:tblStyle')).set(
            _qn('w:val'), self._style("three_line_table"))
        tblW = etree.SubElement(tblPr, _qn('w:tblW'))
        tblW.set(_qn('w:w'), '5000')
        tblW.set(_qn('w:type'), 'pct')

        # Three-line borders
        tblBorders = etree.SubElement(tblPr, _qn('w:tblBorders'))
        for side in ('top', 'bottom'):
            b = etree.SubElement(tblBorders, f'{{{W}}}{side}')
            b.set(_qn('w:val'), 'single')
            b.set(_qn('w:sz'), '12')
            b.set(_qn('w:space'), '0')
            b.set(_qn('w:color'), '000000')
        for side in ('left', 'right', 'insideH', 'insideV'):
            b = etree.SubElement(tblBorders, f'{{{W}}}{side}')
            b.set(_qn('w:val'), 'none')
            b.set(_qn('w:sz'), '0')
            b.set(_qn('w:space'), '0')
            b.set(_qn('w:color'), 'auto')

        etree.SubElement(tblPr, _qn('w:tblLayout')).set(_qn('w:type'), 'fixed')
        tblLook = etree.SubElement(tblPr, _qn('w:tblLook'))
        tblLook.set(_qn('w:val'), '0020')
        tblLook.set(_qn('w:firstRow'), '1')

        # Grid (equal widths)
        page_w = self.profile.get("page", {}).get("content_width_twips", 8306)
        col_w = page_w // n_cols
        tblGrid = etree.SubElement(tbl, _qn('w:tblGrid'))
        for _ in range(n_cols):
            etree.SubElement(tblGrid, _qn('w:gridCol')).set(_qn('w:w'), str(col_w))

        # Cell styling helper: 10.5pt, no indent, center align
        cell_cfg = self.profile.get("table_cell", {})
        cell_sz = str(cell_cfg.get("font_size", 21))  # 10.5pt

        def _make_cell_p(tc, cell_runs, bold=False):
            p = etree.SubElement(tc, _qn('w:p'))
            pPr = etree.SubElement(p, _qn('w:pPr'))
            # 0.25-line spacing before and after each cell paragraph
            sp = etree.SubElement(pPr, _qn('w:spacing'))
            sp.set(_qn('w:beforeLines'), '25')   # 0.25 lines
            sp.set(_qn('w:afterLines'), '25')    # 0.25 lines
            sp.set(_qn('w:line'), '240')
            sp.set(_qn('w:lineRule'), 'auto')
            # No first-line indent, center alignment
            ind = etree.SubElement(pPr, _qn('w:ind'))
            ind.set(_qn('w:firstLine'), '0')
            etree.SubElement(pPr, _qn('w:jc')).set(_qn('w:val'), 'center')
            # Add runs with 10.5pt size override
            for run in cell_runs:
                if run.get("type") == "text":
                    r = etree.SubElement(p, _qn('w:r'))
                    rPr = etree.SubElement(r, _qn('w:rPr'))
                    sz = etree.SubElement(rPr, _qn('w:sz'))
                    sz.set(_qn('w:val'), cell_sz)
                    szCs = etree.SubElement(rPr, _qn('w:szCs'))
                    szCs.set(_qn('w:val'), cell_sz)
                    if bold or run.get("bold"):
                        etree.SubElement(rPr, _qn('w:b'))
                        etree.SubElement(rPr, _qn('w:bCs'))
                    t = etree.SubElement(r, _qn('w:t'))
                    t.text = run.get("text", "")
                    if t.text.startswith(' ') or t.text.endswith(' '):
                        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                elif run.get("type") == "inline_math":
                    self._add_inline_math(p, run)

        # Header row
        if headers:
            tr = etree.SubElement(tbl, _qn('w:tr'))
            trPr = etree.SubElement(tr, _qn('w:trPr'))
            etree.SubElement(trPr, _qn('w:tblHeader')).set(_qn('w:val'), 'on')

            for cell_runs in headers:
                tc = etree.SubElement(tr, _qn('w:tc'))
                tcPr = etree.SubElement(tc, _qn('w:tcPr'))
                # Header bottom border (thin line)
                tcBorders = etree.SubElement(tcPr, _qn('w:tcBorders'))
                hb = etree.SubElement(tcBorders, f'{{{W}}}bottom')
                hb.set(_qn('w:val'), 'single')
                hb.set(_qn('w:sz'), '6')
                hb.set(_qn('w:space'), '0')
                hb.set(_qn('w:color'), '000000')
                etree.SubElement(tcPr, _qn('w:vAlign')).set(_qn('w:val'), 'center')
                _make_cell_p(tc, cell_runs, bold=True)

        # Data rows
        for row in rows:
            tr = etree.SubElement(tbl, _qn('w:tr'))
            for cell_runs in row:
                tc = etree.SubElement(tr, _qn('w:tc'))
                tcPr = etree.SubElement(tc, _qn('w:tcPr'))
                etree.SubElement(tcPr, _qn('w:vAlign')).set(_qn('w:val'), 'center')
                _make_cell_p(tc, cell_runs)

    # ── Section properties ────────────────────────────────────────────

    def _add_section_props(self, body):
        """Add sectPr for page layout + line numbering."""
        page = self.profile.get("page", {})
        sectPr = etree.SubElement(body, _qn('w:sectPr'))
        pgSz = etree.SubElement(sectPr, _qn('w:pgSz'))
        pgSz.set(_qn('w:w'), str(page.get("width", 12240)))
        pgSz.set(_qn('w:h'), str(page.get("height", 15840)))
        pgMar = etree.SubElement(sectPr, _qn('w:pgMar'))
        pgMar.set(_qn('w:top'), str(page.get("margin_top", 1440)))
        pgMar.set(_qn('w:right'), str(page.get("margin_right", 1440)))
        pgMar.set(_qn('w:bottom'), str(page.get("margin_bottom", 1440)))
        pgMar.set(_qn('w:left'), str(page.get("margin_left", 1440)))
        pgMar.set(_qn('w:header'), '432')
        pgMar.set(_qn('w:footer'), '720')
        pgMar.set(_qn('w:gutter'), '0')

        # Line numbering (continuous)
        if page.get("line_numbering"):
            lnNum = etree.SubElement(sectPr, _qn('w:lnNumType'))
            lnNum.set(_qn('w:countBy'), '1')
            lnNum.set(_qn('w:restart'), 'continuous')

        etree.SubElement(sectPr, _qn('w:cols')).set(_qn('w:space'), '720')
        docGrid = etree.SubElement(sectPr, _qn('w:docGrid'))
        docGrid.set(_qn('w:linePitch'), '360')
