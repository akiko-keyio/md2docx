"""Main converter — orchestrates md → docx pipeline.

Usage:
    from md2docx.convert import convert
    convert("input.md", "output.docx", style="academic-manuscript")
"""

import logging
import shutil
from pathlib import Path
from lxml import etree

from .parse import parse
from .build import DocxBuilder, NS
from .math import MathEngine

log = logging.getLogger(__name__)


# ── Style profiles ────────────────────────────────────────────────────
# Each style's visual formatting lives in styles/<name>/ (Word XML).
# Structural parameters are defined here as constants.

_PROFILES = {
    "academic-manuscript": {
        "name": "academic-manuscript",
        "styles": {
            "body_text":        "BodyText",
            "heading1":         "10",
            "heading2":         "20",
            "figure_block":     "FigureBlock",
            "fig_caption":      "FigCaption",
            "equation":         "Equation",
            "equation_table":   "EquationTable",
            "three_line_table": "ThreeLineTable",
        },
        "equation": {
            "col_widths": [8706, 654],
            "spacing_before": 50,
            "spacing_after": 50,
            "number_font": "Times New Roman",
        },
        "caption": {
            "short_threshold": 60,
        },
        "table_caption": {
            "spacing_before_lines": 50,
            "spacing_after_lines": 0,
        },
        "table_cell": {
            "font_size": 21,
            "first_line_indent": 0,
        },
        "page": {
            "width": 12240,
            "height": 15840,
            "margin_top": 1440,
            "margin_bottom": 1440,
            "margin_left": 1440,
            "margin_right": 1440,
            "content_width_cm": 16.51,
            "content_width_twips": 9360,
            "line_numbering": True,
        },
    },
}

# ── Skill root ────────────────────────────────────────────────────────
_SKILL_DIR = Path(__file__).resolve().parent.parent


def convert(md_path: str | Path, output_path: str | Path | None = None,
            style: str = "academic-manuscript") -> Path:
    """Convert a Markdown file to a DOCX file.

    Args:
        md_path: Path to the Markdown source file.
        output_path: Output .docx path. Defaults to same stem as input.
        style: Style preset name (must exist in styles/ directory).

    Returns:
        Path to the generated .docx file.
    """
    md_path = Path(md_path).resolve()
    if output_path is None:
        output_path = md_path.with_suffix('.docx')
    else:
        output_path = Path(output_path).resolve()

    # Resolve style directory
    style_dir = _SKILL_DIR / "styles" / style
    if not style_dir.exists():
        raise FileNotFoundError(
            f"Style not found: {style_dir}\n"
            f"Available styles: {[d.name for d in (_SKILL_DIR / 'styles').iterdir() if d.is_dir()]}"
        )

    # Load profile
    if style not in _PROFILES:
        raise ValueError(f"Unknown style: {style}. Available: {list(_PROFILES.keys())}")
    profile = _PROFILES[style]

    # Parse Markdown
    md_text = md_path.read_text(encoding='utf-8-sig')  # utf-8-sig strips BOM if present
    blocks = parse(md_text, base_dir=md_path.parent)

    log.info("Parsed %d blocks from %s", len(blocks), md_path.name)
    _print_block_summary(blocks)

    # Prepare output directory (copy style template)
    work_dir = output_path.with_suffix('.tmp_unpacked')
    if work_dir.exists():
        shutil.rmtree(work_dir)
    shutil.copytree(style_dir, work_dir)

    # Clean old media and image rels from template
    _clean_template_images(work_dir)

    # Ensure Hyperlink character style exists
    _ensure_hyperlink_style(work_dir)

    # Get max rId from template rels (after cleaning images)
    max_rid = _get_max_rid(work_dir)

    # Build document XML
    math_engine = MathEngine()
    builder = DocxBuilder(profile, math_engine)
    builder._rid_counter = max_rid + 1

    doc_element = builder.build(blocks)

    # Write document.xml
    doc_xml_path = work_dir / 'word' / 'document.xml'
    tree = etree.ElementTree(doc_element)
    tree.write(str(doc_xml_path), xml_declaration=True, encoding='UTF-8', standalone=True)

    # Handle images
    images = builder.images
    if images:
        media_dir = work_dir / 'word' / 'media'
        media_dir.mkdir(exist_ok=True)

        for img_info in images:
            src = Path(img_info["path"])
            dst = media_dir / img_info["media_name"]
            shutil.copy2(src, dst)

        # Update relationships
        _update_relationships(work_dir, images)

        # Update [Content_Types].xml
        _update_content_types(work_dir, images)

    log.info("%d images embedded", len(images))

    # Pack into docx
    _pack_docx(work_dir, output_path)

    # Clean up
    shutil.rmtree(work_dir)

    log.info("Output: %s", output_path)
    return output_path


def _print_block_summary(blocks: list[dict]):
    from collections import Counter
    counts = Counter(b["type"] for b in blocks)
    parts = [f"{v} {k}" for k, v in counts.most_common()]
    log.info("Blocks: %s", ", ".join(parts))


def _get_max_rid(work_dir: Path) -> int:
    """Find the highest rId number in template's document.xml.rels."""
    import re
    rels_path = work_dir / 'word' / '_rels' / 'document.xml.rels'
    if not rels_path.exists():
        return 0
    RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    tree = etree.parse(str(rels_path))
    max_id = 0
    for rel in tree.getroot().findall(f'{{{RELS_NS}}}Relationship'):
        rid = rel.get('Id', '')
        m = re.match(r'rId(\d+)', rid)
        if m:
            max_id = max(max_id, int(m.group(1)))
    return max_id


def _ensure_hyperlink_style(work_dir: Path):
    """Inject the Hyperlink character style into styles.xml if it is missing."""
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    styles_path = work_dir / 'word' / 'styles.xml'
    if not styles_path.exists():
        return

    tree = etree.parse(str(styles_path))
    root = tree.getroot()

    existing = root.findall(f'{{{W}}}style[@{{{W}}}styleId="Hyperlink"]')
    if existing:
        return

    style_xml = f'''<w:style xmlns:w="{W}" w:type="character" w:styleId="Hyperlink">
  <w:name w:val="Hyperlink"/>
  <w:basedOn w:val="DefaultParagraphFont"/>
  <w:uiPriority w:val="99"/>
  <w:unhideWhenUsed/>
  <w:rPr>
    <w:color w:val="0563C1" w:themeColor="hyperlink"/>
    <w:u w:val="single"/>
  </w:rPr>
</w:style>'''
    new_style = etree.fromstring(style_xml)
    root.append(new_style)

    tree.write(str(styles_path), xml_declaration=True, encoding='UTF-8', standalone=True)


def _clean_template_images(work_dir: Path):
    """Remove old image/comment relationships and files from template."""
    RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    REMOVE_TYPES = {
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
        'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
        'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
        'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
        'http://schemas.microsoft.com/office/2011/relationships/people',
    }

    rels_path = work_dir / 'word' / '_rels' / 'document.xml.rels'
    if rels_path.exists():
        tree = etree.parse(str(rels_path))
        root = tree.getroot()
        for rel in root.findall(f'{{{RELS_NS}}}Relationship'):
            if rel.get('Type') in REMOVE_TYPES:
                root.remove(rel)
        tree.write(str(rels_path), xml_declaration=True, encoding='UTF-8', standalone=True)

    # Remove old media directory
    media_dir = work_dir / 'word' / 'media'
    if media_dir.exists():
        shutil.rmtree(media_dir)

    # Remove stale comment XML files
    for name in ('comments.xml', 'commentsExtended.xml',
                 'commentsExtensible.xml', 'commentsIds.xml', 'people.xml'):
        f = work_dir / 'word' / name
        if f.exists():
            f.unlink()

    # Also clean [Content_Types].xml overrides for removed parts
    ct_path = work_dir / '[Content_Types].xml'
    if ct_path.exists():
        CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
        tree = etree.parse(str(ct_path))
        root = tree.getroot()
        for override in root.findall(f'{{{CT_NS}}}Override'):
            part = override.get('PartName', '')
            if any(x in part for x in ('comments', 'people')):
                root.remove(override)
        tree.write(str(ct_path), xml_declaration=True, encoding='UTF-8', standalone=True)


def _update_relationships(work_dir: Path, images: list[dict]):
    """Add image relationships to word/_rels/document.xml.rels."""
    rels_path = work_dir / 'word' / '_rels' / 'document.xml.rels'
    RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

    if rels_path.exists():
        tree = etree.parse(str(rels_path))
        root = tree.getroot()
    else:
        rels_path.parent.mkdir(parents=True, exist_ok=True)
        root = etree.Element(f'{{{RELS_NS}}}Relationships')
        tree = etree.ElementTree(root)

    for img in images:
        rel = etree.SubElement(root, f'{{{RELS_NS}}}Relationship')
        rel.set('Id', img["rId"])
        rel.set('Type', IMAGE_TYPE)
        rel.set('Target', f'media/{img["media_name"]}')

    tree.write(str(rels_path), xml_declaration=True, encoding='UTF-8', standalone=True)


def _update_content_types(work_dir: Path, images: list[dict]):
    """Ensure image extensions are declared in [Content_Types].xml."""
    ct_path = work_dir / '[Content_Types].xml'
    CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

    tree = etree.parse(str(ct_path))
    root = tree.getroot()

    existing_exts = set()
    for default in root.findall(f'{{{CT_NS}}}Default'):
        existing_exts.add(default.get('Extension', '').lower())

    EXT_MIME = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'tiff': 'image/tiff',
        'bmp': 'image/bmp',
    }

    for img in images:
        ext = Path(img["media_name"]).suffix.lstrip('.').lower()
        if ext not in existing_exts and ext in EXT_MIME:
            d = etree.SubElement(root, f'{{{CT_NS}}}Default')
            d.set('Extension', ext)
            d.set('ContentType', EXT_MIME[ext])
            existing_exts.add(ext)

    tree.write(str(ct_path), xml_declaration=True, encoding='UTF-8', standalone=True)


def _pack_docx(src_dir: Path, output_path: Path):
    """Pack directory into a .docx (ZIP) file."""
    import zipfile
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file in sorted(src_dir.rglob('*')):
            if file.is_file():
                arcname = file.relative_to(src_dir).as_posix()
                zf.write(file, arcname)
