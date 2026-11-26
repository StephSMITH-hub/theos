# -*- coding: utf-8 -*-
"""
Script to generate structured clean transcript documents for:
Workers in Training Convention 2026 – What to do with a new convert - Track 1
Speaker: Pastor Chris Segun Onayinka, Saints Community Church
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_para_font(p, font_name="Lexend", font_size=10, bold=False, italic=False, color_rgb=(40, 40, 40)):
    for run in p.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = RGBColor(*color_rgb)

def add_spoken_paragraph(doc, text, font_name="Lexend"):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(40, 40, 40)
    return p

def add_heading_paragraph(doc, text, level=2, font_name="Lexend"):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(12.5 if level == 2 else 11.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(20, 50, 95)
    return p

def add_scripture_header(doc, ref_text, font_name="Lexend"):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(ref_text)
    run.font.name = font_name
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = RGBColor(15, 23, 42)
    return p

def add_scripture_verse(doc, verse_text, font_name="Lexend"):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(verse_text)
    run.font.name = font_name
    run.font.size = Pt(9.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(30, 41, 59)
    return p

print("Builder helper functions loaded.")
