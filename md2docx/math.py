"""LaTeX → OMML conversion engine.

Chain: LaTeX → MathML (latex2mathml) → OMML (Microsoft MML2OMML.xsl)
       → post-process (fix operator names, font)
"""

import re
from pathlib import Path
from lxml import etree
import latex2mathml.converter

# ── Unicode Mathematical Alphanumeric Symbols tables ──────────────────────────
# _STYLED_TO_PLAIN: styled codepoint → (plain_char, style)   ['b' | 'bi']
# _PLAIN_TO_BI:     plain char → bold-italic codepoint
_STYLED_TO_PLAIN: dict[int, tuple[str, str]] = {}
_PLAIN_TO_BI: dict[str, int] = {}


def _init_tables():
    # ── Latin ──
    for i in range(26):
        U, L = chr(ord('A') + i), chr(ord('a') + i)
        _STYLED_TO_PLAIN[0x1D400 + i] = (U, 'b')    # Bold upper
        _STYLED_TO_PLAIN[0x1D41A + i] = (L, 'b')    # Bold lower
        _STYLED_TO_PLAIN[0x1D468 + i] = (U, 'bi')   # Bold-italic upper
        _STYLED_TO_PLAIN[0x1D482 + i] = (L, 'bi')   # Bold-italic lower
        _PLAIN_TO_BI[U] = 0x1D468 + i
        _PLAIN_TO_BI[L] = 0x1D482 + i

    # ── Greek uppercase (25 letters + nabla) ──
    # Position 17 is theta-symbol (ϴ U+03F4), inserted between Rho and Sigma.
    _gu = [
        0x0391, 0x0392, 0x0393, 0x0394, 0x0395, 0x0396, 0x0397, 0x0398,
        0x0399, 0x039A, 0x039B, 0x039C, 0x039D, 0x039E, 0x039F, 0x03A0,
        0x03A1, 0x03F4, 0x03A3, 0x03A4, 0x03A5, 0x03A6, 0x03A7, 0x03A8,
        0x03A9, 0x2207,
    ]
    for i, cp in enumerate(_gu):
        _STYLED_TO_PLAIN[0x1D6A8 + i] = (chr(cp), 'b')
        _STYLED_TO_PLAIN[0x1D71C + i] = (chr(cp), 'bi')
        _PLAIN_TO_BI[chr(cp)] = 0x1D71C + i

    # ── Greek lowercase (25 letters + partial derivative) ──
    _gl = [
        0x03B1, 0x03B2, 0x03B3, 0x03B4, 0x03B5, 0x03B6, 0x03B7, 0x03B8,
        0x03B9, 0x03BA, 0x03BB, 0x03BC, 0x03BD, 0x03BE, 0x03BF, 0x03C0,
        0x03C1, 0x03C2, 0x03C3, 0x03C4, 0x03C5, 0x03C6, 0x03C7, 0x03C8,
        0x03C9, 0x2202,
    ]
    for i, cp in enumerate(_gl):
        _STYLED_TO_PLAIN[0x1D6C2 + i] = (chr(cp), 'b')
        _STYLED_TO_PLAIN[0x1D736 + i] = (chr(cp), 'bi')
        _PLAIN_TO_BI[chr(cp)] = 0x1D736 + i

    # ── Greek symbol variants (epsilon, theta, kappa, phi, rho, pi) ──
    for plain_cp, b_cp, bi_cp in [
        (0x03F5, 0x1D6DC, 0x1D750),   # lunate epsilon ϵ
        (0x03D1, 0x1D6DD, 0x1D751),   # theta symbol   ϑ
        (0x03F0, 0x1D6DE, 0x1D752),   # kappa symbol   ϰ
        (0x03D5, 0x1D6DF, 0x1D753),   # phi symbol     ϕ
        (0x03F1, 0x1D6E0, 0x1D754),   # rho symbol     ϱ
        (0x03D6, 0x1D6E1, 0x1D755),   # pi symbol      ϖ
    ]:
        _STYLED_TO_PLAIN[b_cp] = (chr(plain_cp), 'b')
        _STYLED_TO_PLAIN[bi_cp] = (chr(plain_cp), 'bi')
        _PLAIN_TO_BI[chr(plain_cp)] = bi_cp

    # ── Bold digits 𝟎-𝟗 ──
    for i in range(10):
        _STYLED_TO_PLAIN[0x1D7CE + i] = (chr(ord('0') + i), 'b')


_init_tables()

# LaTeX command name → plain Unicode char (for resolving \boldsymbol{\phi} etc.)
_LATEX_CMD_TO_CHAR = {
    'alpha': '\u03B1', 'beta': '\u03B2', 'gamma': '\u03B3', 'delta': '\u03B4',
    'epsilon': '\u03F5', 'varepsilon': '\u03B5', 'zeta': '\u03B6', 'eta': '\u03B7',
    'theta': '\u03B8', 'vartheta': '\u03D1', 'iota': '\u03B9', 'kappa': '\u03BA',
    'varkappa': '\u03F0', 'lambda': '\u03BB', 'mu': '\u03BC', 'nu': '\u03BD',
    'xi': '\u03BE', 'pi': '\u03C0', 'varpi': '\u03D6', 'rho': '\u03C1',
    'varrho': '\u03F1', 'sigma': '\u03C3', 'tau': '\u03C4', 'upsilon': '\u03C5',
    'phi': '\u03D5', 'varphi': '\u03C6', 'chi': '\u03C7', 'psi': '\u03C8',
    'omega': '\u03C9',
    'Gamma': '\u0393', 'Delta': '\u0394', 'Theta': '\u0398', 'Lambda': '\u039B',
    'Xi': '\u039E', 'Pi': '\u03A0', 'Sigma': '\u03A3', 'Upsilon': '\u03A5',
    'Phi': '\u03A6', 'Psi': '\u03A8', 'Omega': '\u03A9',
    'nabla': '\u2207', 'partial': '\u2202',
}

# Auto-detect MML2OMML.xsl
_CANDIDATES = [
    Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"),
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL"),
    Path(r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL"),
]


def _find_xslt() -> Path:
    for p in _CANDIDATES:
        if p.exists():
            return p
    raise FileNotFoundError(
        "MML2OMML.XSL not found. Install Microsoft Office or specify path manually."
    )


class MathEngine:
    """Convert LaTeX math expressions to OMML XML elements."""

    def __init__(self, xslt_path: str | Path | None = None):
        path = Path(xslt_path) if xslt_path else _find_xslt()
        xslt_doc = etree.parse(str(path))
        self._transform = etree.XSLT(xslt_doc)

    # Known operator/function names that must be upright (not italic)
    _OPERATOR_NAMES = {
        'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
        'arcsin', 'arccos', 'arctan', 'arctan2',
        'sinh', 'cosh', 'tanh', 'coth',
        'log', 'ln', 'exp', 'lim', 'sup', 'inf',
        'min', 'max', 'arg', 'det', 'dim', 'deg',
        'gcd', 'hom', 'ker', 'mod', 'Pr',
        'argmax', 'argmin', 'sgn', 'tr', 'diag',
        'operatorname',
    }

    _LIMIT_OPERATORS = {
        'min', 'max', 'lim', 'sup', 'inf', 'limsup', 'liminf',
    }

    def _preprocess(self, latex: str) -> str:
        """Normalize LaTeX before conversion to avoid round-trip drift.

        `\ ` (backslash-space): latex2mathml encodes trailing backslash-space as a
        visible backslash in OMML; pandoc reads it back as \\backslash.
        Replace with a plain space.  Guard (?<!\\) so we preserve \\ (array row breaks).
        """
        latex = re.sub(r'(?<!\\)\\ ', ' ', latex)
        # \boldsymbol{x} → Unicode math bold-italic characters
        # so latex2mathml passes them through and _postprocess_omml
        # converts them back to <m:sty m:val="bi"/> + plain chars.
        latex = re.sub(r'\\boldsymbol\{([^}]+)\}', self._boldsymbol_to_unicode, latex)
        return latex

    def to_omml(self, latex: str) -> etree._Element:
        """Convert a LaTeX string to an OMML <m:oMath> element."""
        latex = self._preprocess(latex)
        mathml_str = latex2mathml.converter.convert(latex)
        mml_tree = etree.fromstring(mathml_str.encode())
        omml_tree = self._transform(mml_tree)
        root = omml_tree.getroot()
        self._postprocess_omml(root)
        return root

    def to_omml_para(self, latex: str) -> etree._Element:
        """Convert LaTeX to an <m:oMathPara> (display math) element."""
        M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        omath = self.to_omml(latex)

        # Wrap in oMathPara if not already
        if omath.tag == f'{{{M}}}oMathPara':
            return omath

        para = etree.Element(f'{{{M}}}oMathPara')
        para_pr = etree.SubElement(para, f'{{{M}}}oMathParaPr')
        jc = etree.SubElement(para_pr, f'{{{M}}}jc')
        jc.set(f'{{{M}}}val', 'center')
        para.append(omath)
        return para

    def _postprocess_omml(self, root: etree._Element):
        """Fix OMML after XSLT conversion:
        1. Remove explicit fonts — let Word use Cambria Math
        2. Force upright style on operator/function names
        3. Set supHide='on' for nary operators with empty superscript
        4. Convert sSub → limLow for limit-type operators
        5. Replace Unicode math bold/bold-italic chars with OOXML style attrs
        """
        M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        # --- Fix 5: Unicode math styled chars → OOXML style attributes ---
        # Must run before operator-name detection since bold chars won't
        # match _OPERATOR_NAMES until decoded to plain text.
        # Segment each run by style, replace with one <m:r> per segment.
        import copy
        for mr in list(root.iter(f'{{{M}}}r')):
            mt = mr.find(f'{{{M}}}t')
            if mt is None or not mt.text:
                continue
            # Segment text into (decoded_text, style) runs
            segments = []
            cur_text, cur_sty = [], None
            for ch in mt.text:
                entry = _STYLED_TO_PLAIN.get(ord(ch))
                ch_plain, ch_style = (entry[0], entry[1]) if entry else (ch, None)
                if ch_style != cur_sty:
                    if cur_text:
                        segments.append((''.join(cur_text), cur_sty))
                    cur_text, cur_sty = [ch_plain], ch_style
                else:
                    cur_text.append(ch_plain)
            if cur_text:
                segments.append((''.join(cur_text), cur_sty))
            # Skip if single unstyled segment (nothing to fix)
            if len(segments) == 1 and segments[0][1] is None:
                continue
            parent = mr.getparent()
            idx = list(parent).index(mr)
            orig_mrPr = mr.find(f'{{{M}}}rPr')
            parent.remove(mr)
            for si, (seg_text, seg_sty) in enumerate(segments):
                new_mr = etree.Element(f'{{{M}}}r')
                new_mrPr = copy.deepcopy(orig_mrPr) if orig_mrPr is not None else etree.Element(f'{{{M}}}rPr')
                if seg_sty:
                    sty_el = new_mrPr.find(f'{{{M}}}sty')
                    if sty_el is None:
                        sty_el = etree.SubElement(new_mrPr, f'{{{M}}}sty')
                    sty_el.set(f'{{{M}}}val', seg_sty)
                new_mr.append(new_mrPr)
                new_mt = etree.SubElement(new_mr, f'{{{M}}}t')
                new_mt.text = seg_text
                parent.insert(idx + si, new_mr)

        # --- Original: font cleanup + operator names ---
        for mr in root.iter(f'{{{M}}}r'):
            mt = mr.find(f'{{{M}}}t')
            text = mt.text.strip() if mt is not None and mt.text else ''

            # Remove any explicit font — let Word use Cambria Math (default)
            wrPr = mr.find(f'{{{W}}}rPr')
            if wrPr is not None:
                rFonts = wrPr.find(f'{{{W}}}rFonts')
                if rFonts is not None:
                    wrPr.remove(rFonts)

            # Check if this run contains an operator name → force upright
            mrPr = mr.find(f'{{{M}}}rPr')
            if text in self._OPERATOR_NAMES:
                if mrPr is None:
                    mrPr = etree.SubElement(mr, f'{{{M}}}rPr')
                    mr.insert(0, mrPr)
                sty = mrPr.find(f'{{{M}}}sty')
                if sty is None:
                    sty = etree.SubElement(mrPr, f'{{{M}}}sty')
                sty.set(f'{{{M}}}val', 'p')  # 'p' = plain/upright

        # --- Fix 3: supHide for nary with empty superscript ---
        for nary in root.iter(f'{{{M}}}nary'):
            nary_pr = nary.find(f'{{{M}}}naryPr')
            sup_elem = nary.find(f'{{{M}}}sup')
            if sup_elem is not None and len(sup_elem) == 0:
                if not (sup_elem.text and sup_elem.text.strip()):
                    if nary_pr is not None:
                        sup_hide = nary_pr.find(f'{{{M}}}supHide')
                        if sup_hide is None:
                            sup_hide = etree.SubElement(nary_pr, f'{{{M}}}supHide')
                        sup_hide.set(f'{{{M}}}val', 'on')

        # --- Fix 4: sSub → limLow for limit-type operators ---
        for ssub in list(root.iter(f'{{{M}}}sSub')):
            e_elem = ssub.find(f'{{{M}}}e')
            sub_elem = ssub.find(f'{{{M}}}sub')
            if e_elem is None or sub_elem is None:
                continue
            base_texts = [mt.text for mt in e_elem.iter(f'{{{M}}}t') if mt.text]
            base_text = ''.join(base_texts).strip()
            if base_text not in self._LIMIT_OPERATORS:
                continue
            limlow = etree.Element(f'{{{M}}}limLow')
            new_e = etree.SubElement(limlow, f'{{{M}}}e')
            for child in list(e_elem):
                new_e.append(child)
            new_lim = etree.SubElement(limlow, f'{{{M}}}lim')
            for child in list(sub_elem):
                new_lim.append(child)
            parent = ssub.getparent()
            if parent is not None:
                idx = list(parent).index(ssub)
                parent.remove(ssub)
                parent.insert(idx, limlow)

    @staticmethod
    def _decode_unicode_math(text):
        """Decode Unicode math bold/bold-italic chars → (plain_text, style).

        Returns (decoded_text, style_string) where style_string is 'b', 'bi',
        or None if no styled characters were found.
        Uses module-level _STYLED_TO_PLAIN table (Latin + Greek + digits).
        """
        result = []
        style = None
        for ch in text:
            entry = _STYLED_TO_PLAIN.get(ord(ch))
            if entry:
                result.append(entry[0])
                style = entry[1]
            else:
                result.append(ch)
        return ''.join(result), style

    @staticmethod
    def _boldsymbol_to_unicode(m):
        """Convert \\boldsymbol{xyz} match → Unicode math bold-italic chars.

        Handles both plain characters (c → 𝒄) and LaTeX commands
        (\\phi → 𝝓) via _PLAIN_TO_BI + _LATEX_CMD_TO_CHAR tables.
        """
        content = m.group(1)
        # Single LaTeX command: \phi, \Phi, \alpha, ...
        cmd_m = re.match(r'^\\([a-zA-Z]+)$', content)
        if cmd_m:
            plain_char = _LATEX_CMD_TO_CHAR.get(cmd_m.group(1))
            if plain_char:
                bi_cp = _PLAIN_TO_BI.get(plain_char)
                if bi_cp is not None:
                    return chr(bi_cp)
            # Unknown command — return original \boldsymbol{...} intact
            return m.group(0)
        # Plain characters: encode each individually
        result = []
        for ch in content:
            bi_cp = _PLAIN_TO_BI.get(ch)
            if bi_cp is not None:
                result.append(chr(bi_cp))
            else:
                result.append(ch)
        return ''.join(result)
