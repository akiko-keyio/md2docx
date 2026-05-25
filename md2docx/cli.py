"""CLI entry point for md2docx converter.

Usage:
    python -m md2docx input.md [-o output.docx] [--style academic-manuscript]
"""

import argparse
import logging
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(
        description='Convert Markdown to Word (.docx).')
    parser.add_argument('input', help='Input Markdown file')
    parser.add_argument('-o', '--output', help='Output .docx file (default: same name as input)')
    parser.add_argument('--style', default='academic-manuscript',
                        help='Style preset name (default: academic-manuscript)')

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format='%(levelname)s: %(message)s',
        stream=sys.stdout,
    )

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        print(f"Error: {input_path} not found", file=sys.stderr)
        sys.exit(1)

    output_path = Path(args.output).resolve() if args.output else None

    from .convert import convert
    convert(input_path, output_path, args.style)


if __name__ == '__main__':
    main()
