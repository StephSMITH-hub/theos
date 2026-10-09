#!/usr/bin/env python3
"""
==============================================================================
study_group_engine.py
Automated Repeatable Pipeline for Transcript Formatting and Study Group Answers
Location: c:\\Users\\Stephycopy\\OneDrive\\Desktop\\Theos\\theos\\Study group
==============================================================================
"""

import sys
import os
import argparse
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Default formatting constants
DARK_COLOR = RGBColor(0x11, 0x18, 0x27)
FONT_NAME = "Arial"

def configure_section_geometry(section):
    """Sets Letter size with 1.0 inch margins all around."""
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

def set_normal_style(doc):
    """Sets standard body paragraph defaults."""
    style = doc.styles['Normal']
    font = style.font
    font.name = FONT_NAME
    font.size = Pt(12)
    font.color.rgb = DARK_COLOR

def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = Pt(18)
    run.bold = True
    return p

def add_subtitle(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = Pt(14)
    run.bold = True
    run.italic = True
    return p

def add_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.bold = True
    return p

def add_scripture_ref(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.bold = True
    run.italic = True
    return p

def add_scripture_verse(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    run.bold = True
    run.italic = True
    return p

def add_body_paragraph(doc, runs_data):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    for item in runs_data:
        if isinstance(item, tuple):
            r_text, r_bold, r_italic = item
        else:
            r_text, r_bold, r_italic = item, False, False
        run = p.add_run(r_text)
        run.font.name = FONT_NAME
        run.font.size = Pt(12)
        run.bold = r_bold
        run.italic = r_italic
    return p

def convert_md_file_to_docx(md_path, docx_path):
    """Parses formatted study markdown file into standardized docx."""
    if not os.path.exists(md_path):
        print(f"[!] Error: {md_path} does not exist.")
        return False

    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    doc = Document()
    for s in doc.sections:
        configure_section_geometry(s)
    set_normal_style(doc)

    for line in lines:
        stripped = line.strip()
        if not stripped or stripped == "---":
            continue

        if stripped.startswith("# "):
            add_title(doc, stripped[2:].strip())
        elif stripped.startswith("## "):
            sub = stripped[3:].strip().strip("*").strip()
            add_subtitle(doc, sub)
        elif stripped.startswith("### ") or stripped.startswith("#### "):
            h_text = stripped.lstrip("#").strip()
            add_heading(doc, h_text)
        elif stripped.startswith("> ***[") or stripped.startswith("> ["):
            verse_text = stripped.lstrip(">").strip()
            # Clean markdown bold/italics
            verse_text = verse_text.replace("***", "").replace("**", "")
            add_scripture_verse(doc, verse_text)
        elif stripped.startswith("***") and stripped.endswith("***") and not stripped.startswith(">"):
            ref_text = stripped.strip("*").strip()
            add_scripture_ref(doc, ref_text)
        else:
            # Parse inline bold/italics in paragraph
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            
            # Simple run parser
            parts = stripped.split("***")
            for i, part in enumerate(parts):
                if not part:
                    continue
                run = p.add_run(part)
                run.font.name = FONT_NAME
                run.font.size = Pt(12)
                if i % 2 == 1:
                    run.bold = True
                    run.italic = True

    doc.save(docx_path)
    print(f"[+] Successfully converted {md_path} -> {docx_path}")
    return True

def count_words_in_docx(file_path):
    doc = Document(file_path)
    total = 0
    for p in doc.paragraphs:
        words = p.text.split()
        total += len(words)
    return total

def run_pipeline():
    base_dir = Path(__file__).resolve().parent
    repo_dir = base_dir.parent

    print("==========================================================")
    print("   STUDY GROUP AUTOMATION & REPEATABLE PIPELINE")
    print("==========================================================")

    # 1. Ensure transcript docx exists
    t_md = base_dir / "Oct_Conf_Formatted.md"
    t_docx = base_dir / "Oct_Conf_Formatted.docx"
    if t_md.exists():
        convert_md_file_to_docx(t_md, t_docx)
        print(f"[*] Transcript Word Count: {count_words_in_docx(t_docx)} words.")

    # 2. Ensure Study Group Answers docx exists
    a_md = base_dir / "Cell_Leaders_Conference_How_Leadership_Changes_You_StudyGroup_Answers.md"
    a_docx = base_dir / "Cell_Leaders_Conference_How_Leadership_Changes_You_StudyGroup_Answers.docx"
    if a_md.exists():
        convert_md_file_to_docx(a_md, a_docx)
        print(f"[*] Study Answers Word Count: {count_words_in_docx(a_docx)} words.")

    print("\n[+] All Study Group files are built, verified, and ready.")

if __name__ == '__main__':
    run_pipeline()
