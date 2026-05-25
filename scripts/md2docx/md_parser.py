"""Markdown parser — converts Markdown text into structured document blocks.

Block types:
  heading       {level, runs}
  paragraph     {runs}
  display_math  {latex, tag}
  figure        {path, width_hint}
  caption       {runs, ref_type, ref_num}
  table         {headers, rows}  (headers/rows are lists of cell run-lists)

Run types (inside runs lists):
  text          {text, bold, italic}
  inline_math   {latex}
  softbreak     {}
"""

import re
from pathlib import Path
from markdown_it import MarkdownIt
from mdit_py_plugins.dollarmath import dollarmath_plugin


def _init_parser() -> MarkdownIt:
    md = MarkdownIt("commonmark", {"breaks": False, "typographer": True})
    md.enable("table")
    md.enable("replacements")  # activates -- → en-dash, --- → em-dash
    dollarmath_plugin(md, double_inline=True)
    return md


_TAG_RE = re.compile(r'\\tag\{(\d+)\}')
_QQUAD_NUM_RE = re.compile(r'\\qquad\s*\((\d+)\)')
_CAPTION_RE = re.compile(r'^(Fig\.\s*\d+|Table\s*\d+\.)')


def _extract_tag(latex: str) -> tuple[str, str | None]:
    """Strip \\tag{N} or \\qquad (N) from LaTeX, return (cleaned, tag_num)."""
    m = _TAG_RE.search(latex)
    if m:
        return _TAG_RE.sub('', latex).strip(), m.group(1)
    m = _QQUAD_NUM_RE.search(latex)
    if m:
        return _QQUAD_NUM_RE.sub('', latex).strip(), m.group(1)
    # Also match bare trailing (N) like "... \qquad (9)"
    m = re.search(r'\((\d+)\)\s*$', latex)
    if m:
        return latex[:m.start()].rstrip(), m.group(1)
    return latex, None


def _parse_inline_tokens(tokens: list) -> list[dict]:
    """Convert markdown-it inline children to a list of run dicts."""
    runs = []
    bold = False
    italic = False

    for tok in tokens:
        if tok.type == 'strong_open':
            bold = True
        elif tok.type == 'strong_close':
            bold = False
        elif tok.type == 'em_open':
            italic = True
        elif tok.type == 'em_close':
            italic = False
        elif tok.type == 'text':
            runs.append({"type": "text", "text": tok.content, "bold": bold, "italic": italic})
        elif tok.type == 'code_inline':
            runs.append({"type": "text", "text": tok.content, "bold": bold, "italic": True})
        elif tok.type == 'math_inline':
            runs.append({"type": "inline_math", "latex": tok.content})
        elif tok.type == 'softbreak':
            runs.append({"type": "text", "text": " ", "bold": bold, "italic": italic})
        elif tok.type == 'hardbreak':
            runs.append({"type": "break"})
        elif tok.type == 'image':
            # Image inside inline — extract src and alt text
            src = tok.attrGet('src') or ''
            runs.append({"type": "image", "src": src, "alt": tok.content or ''})

    return runs


def _is_caption(runs: list[dict]) -> tuple[bool, str | None, str | None]:
    """Check if runs form a figure/table caption. Returns (is_cap, ref_type, ref_num)."""
    if not runs:
        return False, None, None
    # Skip leading empty text runs
    first = None
    for r in runs:
        if r.get("type") == "text" and r.get("text", "").strip() == "":
            continue
        first = r
        break
    if first is None:
        return False, None, None
    if first.get("type") != "text" or not first.get("bold"):
        return False, None, None
    text = first["text"].strip()
    m_fig = re.match(r'^Fig\.\s*(\d+)', text)
    if m_fig:
        return True, "fig", m_fig.group(1)
    m_tbl = re.match(r'^Table\s*(\d+)\.', text)
    if m_tbl:
        return True, "table", m_tbl.group(1)
    return False, None, None


HALF_WIDTH_THRESHOLD_PX = 1500  # images ≤ this pixel width → half column


def _detect_width_hint(img_path: str) -> str:
    """Detect whether an image should be half or full column width."""
    try:
        from PIL import Image
        with Image.open(img_path) as im:
            w, _ = im.size
            return "half" if w <= HALF_WIDTH_THRESHOLD_PX else "full"
    except Exception:
        return "full"


def parse(md_text: str, base_dir: Path | None = None) -> list[dict]:
    """Parse Markdown text into a list of document blocks."""
    parser = _init_parser()
    tokens = parser.parse(md_text)

    blocks: list[dict] = []
    i = 0

    while i < len(tokens):
        tok = tokens[i]

        # --- Heading ---
        if tok.type == 'heading_open':
            level = int(tok.tag[1])  # h1 → 1, h2 → 2
            i += 1  # inline token
            inline_tok = tokens[i]
            runs = _parse_inline_tokens(inline_tok.children or [])
            blocks.append({"type": "heading", "level": level, "runs": runs})
            i += 2  # skip heading_close
            continue

        # --- Display math ---
        if tok.type == 'math_block' or tok.type == 'math_block_double':
            latex, tag = _extract_tag(tok.content.strip())
            blocks.append({"type": "display_math", "latex": latex, "tag": tag})
            i += 1
            continue

        # --- Table ---
        if tok.type == 'table_open':
            headers = []
            rows = []
            i += 1
            in_head = False
            in_body = False
            current_row = []

            while i < len(tokens) and tokens[i].type != 'table_close':
                t = tokens[i]
                if t.type == 'thead_open':
                    in_head = True
                elif t.type == 'thead_close':
                    in_head = False
                elif t.type == 'tbody_open':
                    in_body = True
                elif t.type == 'tbody_close':
                    in_body = False
                elif t.type == 'tr_open':
                    current_row = []
                elif t.type == 'tr_close':
                    if in_head:
                        headers = current_row
                    else:
                        rows.append(current_row)
                elif t.type in ('th_open', 'td_open'):
                    pass
                elif t.type == 'inline':
                    cell_runs = _parse_inline_tokens(t.children or [])
                    current_row.append(cell_runs)
                elif t.type in ('th_close', 'td_close'):
                    pass
                i += 1

            blocks.append({"type": "table", "headers": headers, "rows": rows})
            i += 1  # skip table_close
            continue

        # --- Paragraph (may contain images, captions, or regular text) ---
        if tok.type == 'paragraph_open':
            i += 1
            inline_tok = tokens[i]
            runs = _parse_inline_tokens(inline_tok.children or [])

            # Check if this is a standalone image
            image_runs = [r for r in runs if r.get("type") == "image"]
            if image_runs and len(runs) == len(image_runs):
                for img in image_runs:
                    src = img["src"]
                    alt = (img.get("alt") or "").strip().lower()
                    if base_dir:
                        src = str((base_dir / src).resolve())
                    if alt in ("half", "full"):
                        width_hint = alt          # explicit override
                    else:
                        width_hint = _detect_width_hint(src)
                    blocks.append({"type": "figure", "path": src, "width_hint": width_hint})
            else:
                # Check if this is a caption
                is_cap, ref_type, ref_num = _is_caption(runs)
                if is_cap:
                    blocks.append({
                        "type": "caption",
                        "runs": runs,
                        "ref_type": ref_type,
                        "ref_num": ref_num,
                    })
                else:
                    blocks.append({"type": "paragraph", "runs": runs})

            i += 2  # skip paragraph_close
            continue

        # --- Fallback: skip unknown tokens ---
        i += 1

    return blocks
