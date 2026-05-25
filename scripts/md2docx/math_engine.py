"""LaTeX → OMML conversion engine.

Chain: LaTeX → MathML (latex2mathml) → OMML (Microsoft MML2OMML.xsl)
       → post-process (fix operator names, font)
"""

import re
from pathlib import Path
from lxml import etree
import latex2mathml.converter

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
        'arcsin', 'arccos', 'arctan', 'arctan2', 'argmax', 'argmin',
        'sinh', 'cosh', 'tanh', 'coth',
        'log', 'ln', 'exp', 'lim', 'sup', 'inf',
        'min', 'max', 'arg', 'det', 'dim', 'deg',
        'gcd', 'hom', 'ker', 'mod', 'Pr',
        'ZTD', 'RMS', 'STD', 'RMSE',
        'subject', 'to',
    }

    def to_omml(self, latex: str) -> etree._Element:
        """Convert a LaTeX string to an OMML <m:oMath> element."""
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
        1. Set XITS Math font on all <m:r> runs
        2. Force upright style on operator/function names
        """
        M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        for mr in root.iter(f'{{{M}}}r'):
            # Get text content
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
