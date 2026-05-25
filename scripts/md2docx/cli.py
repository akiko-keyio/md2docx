"""CLI entry point for md2docx converter.

Usage:
    python -m scripts.md2docx input.md [-o output.docx] [-p academic-manuscript]
"""

import argparse
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(
        description='Convert Markdown to Word (.docx) with configurable style profiles.')
    parser.add_argument('input', help='Input Markdown file')
    parser.add_argument('-o', '--output', help='Output .docx file (default: same name as input)')
    parser.add_argument('-p', '--profile', default='academic-manuscript',
                        help='Style profile (default: academic-manuscript)')
    parser.add_argument('--project-root', help='Project root directory')

    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        print(f"Error: {input_path} not found", file=sys.stderr)
        sys.exit(1)

    output_path = Path(args.output).resolve() if args.output else None
    project_root = Path(args.project_root).resolve() if args.project_root else None

    from .converter import convert
    convert(input_path, output_path, args.profile, project_root)


if __name__ == '__main__':
    main()
