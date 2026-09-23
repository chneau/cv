#!/usr/bin/env -S uv run
# /// script
# requires-python = ">=3.14"
# dependencies = ["python-docx"]
# ///

"""
CLI tool to compile and manage CV formats (Markdown -> DOCX, PDF).
Usage:
    ./build_cv.py              # Build both DOCX and PDF (default)
    ./build_cv.py all          # Build both DOCX and PDF
    ./build_cv.py docx         # Build DOCX only
    ./build_cv.py pdf          # Build PDF only
    ./build_cv.py clean        # Clean generated DOCX and PDF files
    ./build_cv.py watch        # Watch cv.md and auto-rebuild on change
"""

import argparse
import os
import re
import subprocess
import sys
import time
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


def build_docx(md_path: Path, docx_path: Path):
    doc = Document()

    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    # Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(9.5)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    normal_style.paragraph_format.space_after = Pt(3)
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.line_spacing = 1.12

    # Parse markdown content
    content = md_path.read_text(encoding="utf-8")
    lines = content.splitlines()

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        # Header Title (# Charles Neau)
        if line.startswith("# "):
            name = line[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(name)
            run.bold = True
            run.font.size = Pt(20)
            run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
            i += 1
            continue

        # Subtitle / Role (**Senior / Lead Software Engineer & Systems Architect**)
        if line.startswith("**") and ("Lead" in line or "Engineer" in line or "Architect" in line):
            role = line.strip("*")
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(role)
            run.bold = True
            run.font.size = Pt(11.5)
            run.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
            i += 1
            continue

        # Contact info row (Aberdeen, UK • ...)
        if "Aberdeen, UK" in line or "chneau.github.io" in line:
            clean_text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', line)
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(4)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_text)
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            i += 1
            continue

        # Divider (---)
        if line == "---":
            i += 1
            continue

        # Section Header (## Professional Summary, etc.)
        if line.startswith("## "):
            sec_title = line[3:].strip().upper()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(sec_title)
            run.bold = True
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0x09, 0x69, 0xDA)
            
            pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="0969DA"/></w:pBdr>')
            p._p.get_or_add_pPr().append(pBdr)
            i += 1
            continue

        # Job Title (### **Lead Software Engineer & Systems Architect**)
        if line.startswith("### "):
            job_title = re.sub(r'[#*]', '', line).strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(1)
            run = p.add_run(job_title)
            run.bold = True
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
            i += 1
            continue

        # Company / Dates (*Celerum Ltd & PlanSea Solutions* — Aberdeen, UK)
        if line.startswith("*") and line.endswith("*") and ("—" in line or "–" in line or "-" in line):
            comp_dates = line.strip("*")
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(3)
            run = p.add_run(comp_dates)
            run.italic = True
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
            i += 1
            continue

        # Bullet Items (* **Title:** Description or * Description)
        if line.startswith("* "):
            bullet_text = line[2:].strip()
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.12

            bold_match = re.match(r'^\*\*([^*]+)\*\*(.*)$', bullet_text)
            if bold_match:
                prefix, rest = bold_match.groups()
                run_b = p.add_run(prefix)
                run_b.bold = True
                run_b.font.size = Pt(9.5)
                run_b.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

                clean_rest = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', rest)
                clean_rest = re.sub(r'\*\*([^*]+)\*\*', r'\1', clean_rest)
                clean_rest = re.sub(r'\*([^*]+)\*', r'\1', clean_rest)
                run_r = p.add_run(clean_rest)
                run_r.font.size = Pt(9.5)
            else:
                clean_text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', bullet_text)
                clean_text = re.sub(r'\*\*([^*]+)\*\*', r'\1', clean_text)
                clean_text = re.sub(r'\*([^*]+)\*', r'\1', clean_text)
                run = p.add_run(clean_text)
                run.font.size = Pt(9.5)
            i += 1
            continue

        # Standard Paragraph
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        clean_text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', line)
        clean_text = re.sub(r'\*\*([^*]+)\*\*', r'\1', clean_text)
        clean_text = re.sub(r'\*([^*]+)\*', r'\1', clean_text)
        run = p.add_run(clean_text)
        run.font.size = Pt(9.5)
        i += 1

    doc.save(str(docx_path))
    print(f"✓ Generated DOCX: {docx_path.name}")


def build_pdf(md_path: Path, pdf_path: Path):
    cmd = [
        "pandoc",
        str(md_path),
        "-o", str(pdf_path),
        "--pdf-engine=typst",
        "-V", "margin-x=1.5cm",
        "-V", "margin-y=1.5cm",
        "-V", "fontsize=10pt",
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Error compiling PDF with pandoc/typst: {res.stderr}", file=sys.stderr)
        sys.exit(res.returncode)
    print(f"✓ Generated PDF:  {pdf_path.name}")


def clean_artifacts(docx_path: Path, pdf_path: Path):
    removed = 0
    for p in [docx_path, pdf_path]:
        if p.exists():
            p.unlink()
            print(f"Removed: {p.name}")
            removed += 1
    if removed == 0:
        print("No generated artifacts to clean.")


def watch_file(md_path: Path, docx_path: Path, pdf_path: Path):
    print(f"👀 Watching '{md_path.name}' for changes... (Press Ctrl+C to stop)")
    last_mtime = md_path.stat().st_mtime
    # Initial build
    build_docx(md_path, docx_path)
    build_pdf(md_path, pdf_path)

    try:
        while True:
            time.sleep(1)
            if md_path.exists():
                mtime = md_path.stat().st_mtime
                if mtime != last_mtime:
                    last_mtime = mtime
                    print(f"\n[!] Change detected in {md_path.name} at {time.strftime('%H:%M:%S')}")
                    build_docx(md_path, docx_path)
                    build_pdf(md_path, pdf_path)
    except KeyboardInterrupt:
        print("\nStopped watching.")


def main():
    parser = argparse.ArgumentParser(
        prog="build_cv.py",
        description="CLI tool to build and manage CV formats from Markdown."
    )
    parser.add_argument(
        "action",
        nargs="?",
        default="all",
        choices=["all", "docx", "pdf", "clean", "watch"],
        help="Action to perform: all (default), docx, pdf, clean, watch"
    )
    parser.add_argument(
        "--input", "-i",
        type=Path,
        default=Path("cv.md"),
        help="Input markdown file (default: cv.md)"
    )

    args = parser.parse_args()
    md_file = args.input.resolve()

    docx_out = md_file.parent / f"{md_file.stem}.docx"
    pdf_out = md_file.parent / f"{md_file.stem}.pdf"

    if args.action == "clean":
        clean_artifacts(docx_out, pdf_out)
        return

    if not md_file.exists():
        print(f"Error: input file '{md_file}' does not exist.", file=sys.stderr)
        sys.exit(1)

    if args.action == "watch":
        watch_file(md_file, docx_out, pdf_out)
        return

    if args.action in ("all", "docx"):
        build_docx(md_file, docx_out)

    if args.action in ("all", "pdf"):
        build_pdf(md_file, pdf_out)


if __name__ == "__main__":
    main()
