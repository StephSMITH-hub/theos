#!/usr/bin/env python3
"""
==============================================================================
convert_md_to_docx.py
Converts Markdown (.md) documents to Microsoft Word (.docx) files locally.
Tailored for theological study materials, transcripts, sermon series, and books.
Repository: https://github.com/StephSMITH-hub/theos
==============================================================================
"""

import os
import sys
import re
import argparse
import subprocess
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime

# ----------------------------------------------------------------------------
# 1. Dependency Auto-Installation (python-docx)
# ----------------------------------------------------------------------------
HAVE_DOCX = False
try:
    import docx
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    HAVE_DOCX = True
except ImportError:
    print("[*] 'python-docx' library not found. Attempting automatic installation...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
        import docx
        from docx import Document
        from docx.shared import Inches, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.oxml import OxmlElement, parse_xml
        from docx.oxml.ns import nsdecls, qn
        HAVE_DOCX = True
        print("[+] Successfully installed 'python-docx'.")
    except Exception as e:
        print(f"[!] Warning: Could not install 'python-docx' automatically ({e}).")
        print("[*] Falling back to built-in pure OpenXML engine.")
        HAVE_DOCX = False


# ----------------------------------------------------------------------------
# 2. Markdown Inline Parser (Tokens: Bold, Italic, Code, Scripture refs)
# ----------------------------------------------------------------------------
def parse_inline_markdown(text):
    """
    Parses inline markdown syntax and returns a list of token dicts:
    [{'text': '...', 'bold': bool, 'italic': bool, 'code': bool, 'link': str|None}]
    """
    if not text:
        return []

    # Clean up escaped quotes frequently found in transcripts
    text = text.replace(r'\"', '"').replace(r"\'", "'")

    tokens = []
    # Regex to match code, bold-italic, bold, italic, and links
    pattern = re.compile(
        r'(`(?P<code>.+?)`)'
        r'|(\*\*\*(?P<bold_italic>.+?)\*\*\*)'
        r'|(___(?P<bold_italic2>.+?)___)'
        r'|(\*\*(?P<bold>.+?)\*\*)'
        r'|(__(?P<bold2>.+?)__)'
        r'|(\*(?P<italic>[^*]+?)\*)'
        r'|(_(?P<italic2>[^_]+?)_)'
        r'|(\[(?P<link_text>[^\]]+)\]\((?P<link_url>[^\)]+)\))'
    )

    last_idx = 0
    for match in pattern.finditer(text):
        start, end = match.span()
        if start > last_idx:
            tokens.append({
                'text': text[last_idx:start],
                'bold': False,
                'italic': False,
                'code': False,
                'link': None
            })

        gd = match.groupdict()
        if gd.get('code'):
            tokens.append({'text': gd['code'], 'bold': False, 'italic': False, 'code': True, 'link': None})
        elif gd.get('bold_italic') or gd.get('bold_italic2'):
            t = gd.get('bold_italic') or gd.get('bold_italic2')
            tokens.append({'text': t, 'bold': True, 'italic': True, 'code': False, 'link': None})
        elif gd.get('bold') or gd.get('bold2'):
            t = gd.get('bold') or gd.get('bold2')
            tokens.append({'text': t, 'bold': True, 'italic': False, 'code': False, 'link': None})
        elif gd.get('italic') or gd.get('italic2'):
            t = gd.get('italic') or gd.get('italic2')
            tokens.append({'text': t, 'bold': False, 'italic': True, 'code': False, 'link': None})
        elif gd.get('link_text'):
            tokens.append({'text': gd['link_text'], 'bold': False, 'italic': False, 'code': False, 'link': gd['link_url']})

        last_idx = end

    if last_idx < len(text):
        tokens.append({
            'text': text[last_idx:],
            'bold': False,
            'italic': False,
            'code': False,
            'link': None
        })

    return tokens


# ----------------------------------------------------------------------------
# 3. python-docx Based High-Quality Converter
# ----------------------------------------------------------------------------
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set inner cell padding in twips."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_shading(cell, color_hex="F2F4F7"):
    """Set background fill color for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)


def add_runs_to_paragraph(paragraph, tokens, default_font="Calibri", default_size_pt=11, default_color=None):
    """Applies parsed inline tokens to a Word paragraph with styling."""
    for token in tokens:
        run = paragraph.add_run(token['text'])
        run.font.name = default_font
        run.font.size = Pt(default_size_pt)
        if default_color:
            run.font.color.rgb = default_color

        if token['bold']:
            run.bold = True
        if token['italic']:
            run.italic = True
        if token['code']:
            run.font.name = "Consolas"
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(160, 40, 40)
        if token['link']:
            run.font.color.rgb = RGBColor(26, 82, 118)
            run.underline = True


def convert_with_docx(md_path, docx_path, font_family="Calibri"):
    """Converts markdown to docx using python-docx with professional styling."""
    doc = Document()

    # Set 1-inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.different_first_page_header_footer = False

    # Palette Configuration
    COLOR_TITLE = RGBColor(15, 44, 89)       # Deep Navy
    COLOR_H1 = RGBColor(27, 54, 93)          # Navy Blue
    COLOR_H2 = RGBColor(44, 94, 138)         # Medium Navy
    COLOR_H3 = RGBColor(59, 105, 152)        # Slate Blue
    COLOR_BODY = RGBColor(34, 34, 34)        # Off-black / Charcoal
    COLOR_QUOTE = RGBColor(60, 64, 67)       # Slate Gray
    COLOR_SCRIPTURE = RGBColor(20, 60, 100)  # Theological Scripture Tint

    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    in_code_block = False
    code_block_lines = []
    table_lines = []

    def flush_table(tbl_lines):
        if not tbl_lines:
            return
        parsed_rows = []
        for tl in tbl_lines:
            cells = [c.strip() for c in tl.strip().strip('|').split('|')]
            # Skip markdown separator row like |---|---|
            if all(re.match(r'^:?-+:?$', c) for c in cells if c):
                continue
            parsed_rows.append(cells)

        if not parsed_rows:
            return

        max_cols = max(len(r) for r in parsed_rows)
        table = doc.add_table(rows=len(parsed_rows), cols=max_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True

        for row_idx, row_data in enumerate(parsed_rows):
            is_header = (row_idx == 0)
            row = table.rows[row_idx]
            for col_idx in range(max_cols):
                cell = row.cells[col_idx]
                text = row_data[col_idx] if col_idx < len(row_data) else ""
                set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
                if is_header:
                    set_cell_shading(cell, "EBF3FA")
                else:
                    if row_idx % 2 == 1:
                        set_cell_shading(cell, "F9FBFC")
                    else:
                        set_cell_shading(cell, "FFFFFF")

                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                tokens = parse_inline_markdown(text)
                add_runs_to_paragraph(
                    p, tokens, default_font=font_family,
                    default_size_pt=10 if not is_header else 10.5,
                    default_color=COLOR_H1 if is_header else COLOR_BODY
                )
                if is_header:
                    for r in p.runs:
                        r.bold = True

        # Add small spacing after table
        after_p = doc.add_paragraph()
        after_p.paragraph_format.space_before = Pt(4)
        after_p.paragraph_format.space_after = Pt(6)

    idx = 0
    total_lines = len(lines)
    while idx < total_lines:
        line = lines[idx]
        stripped = line.strip()

        # Handle Code Blocks ```
        if stripped.startswith("```"):
            if in_code_block:
                in_code_block = False
                code_text = "\n".join(code_block_lines)
                code_block_lines = []
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.4)
                p.paragraph_format.right_indent = Inches(0.4)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(8)
                run = p.add_run(code_text)
                run.font.name = "Consolas"
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(50, 50, 50)
            else:
                in_code_block = True
                code_block_lines = []
            idx += 1
            continue

        if in_code_block:
            code_block_lines.append(line.rstrip("\r\n"))
            idx += 1
            continue

        # Handle Markdown Tables
        if stripped.startswith("|") and stripped.endswith("|"):
            table_lines.append(stripped)
            idx += 1
            continue
        elif table_lines:
            flush_table(table_lines)
            table_lines = []

        # Empty lines
        if not stripped:
            idx += 1
            continue

        # Horizontal Rule
        if re.match(r'^(\-{3,}|\*{3,}|_{3,})$', stripped):
            hr = doc.add_paragraph()
            hr.paragraph_format.space_before = Pt(10)
            hr.paragraph_format.space_after = Pt(10)
            r = hr.add_run("―" * 45)
            r.font.name = font_family
            r.font.color.rgb = RGBColor(200, 205, 210)
            hr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            idx += 1
            continue

        # Headings (# H1, ## H2, ### H3, #### H4)
        h_match = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if h_match:
            level = len(h_match.group(1))
            h_text = h_match.group(2).strip()

            if level == 1:
                p = doc.add_heading(level=1)
                p.paragraph_format.space_before = Pt(14)
                p.paragraph_format.space_after = Pt(8)
                tokens = parse_inline_markdown(h_text)
                add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=18, default_color=COLOR_TITLE)
                for r in p.runs:
                    r.bold = True
            elif level == 2:
                p = doc.add_heading(level=2)
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(6)
                tokens = parse_inline_markdown(h_text)
                add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=14, default_color=COLOR_H1)
                for r in p.runs:
                    r.bold = True
            elif level == 3:
                p = doc.add_heading(level=3)
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(4)
                tokens = parse_inline_markdown(h_text)
                add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=12.5, default_color=COLOR_H2)
                for r in p.runs:
                    r.bold = True
            else:
                p = doc.add_heading(level=4)
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(3)
                tokens = parse_inline_markdown(h_text)
                add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=11.5, default_color=COLOR_H3)
                for r in p.runs:
                    r.bold = True
            idx += 1
            continue

        # Blockquote (> text) or Scripture Verse Blocks
        if stripped.startswith(">"):
            quote_text = re.sub(r'^>\s?', '', stripped)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.45)
            p.paragraph_format.right_indent = Inches(0.3)
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(6)
            tokens = parse_inline_markdown(quote_text)
            add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=10.5, default_color=COLOR_QUOTE)
            for r in p.runs:
                r.italic = True
            idx += 1
            continue

        # Scripture Verse Citation detection (e.g., "[16] And he that stealeth...")
        verse_match = re.match(r'^(\[\d+\])\s*(.*)$', stripped)
        if verse_match:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.35)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)

            # Add verse number in bold superscript/small
            num_run = p.add_run(verse_match.group(1) + " ")
            num_run.font.name = font_family
            num_run.font.size = Pt(10)
            num_run.bold = True
            num_run.font.color.rgb = COLOR_H1

            # Verse body
            tokens = parse_inline_markdown(verse_match.group(2))
            add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=10.5, default_color=COLOR_SCRIPTURE)
            for r in p.runs[1:]:
                r.italic = True
            idx += 1
            continue

        # Bullet Lists (- item, * item, + item)
        bullet_match = re.match(r'^[-*+]\s+(.*)$', stripped)
        if bullet_match:
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            tokens = parse_inline_markdown(bullet_match.group(1))
            add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=11, default_color=COLOR_BODY)
            idx += 1
            continue

        # Numbered Lists (1. item, 2. item)
        num_list_match = re.match(r'^\d+\.\s+(.*)$', stripped)
        if num_list_match:
            p = doc.add_paragraph(style='List Number')
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            tokens = parse_inline_markdown(num_list_match.group(1))
            add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=11, default_color=COLOR_BODY)
            idx += 1
            continue

        # Standard Paragraph
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        tokens = parse_inline_markdown(stripped)
        add_runs_to_paragraph(p, tokens, default_font=font_family, default_size_pt=11, default_color=COLOR_BODY)

        idx += 1

    # Final table flush if document ended with a table
    if table_lines:
        flush_table(table_lines)

    doc.save(docx_path)


# ----------------------------------------------------------------------------
# 4. Pure Standard-Library OpenXML Fallback Engine (Zero external dependencies)
# ----------------------------------------------------------------------------
def escape_xml(s):
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;")
             .replace("'", "&apos;"))


def convert_with_openxml_fallback(md_path, docx_path):
    """
    Creates a valid Microsoft Word (.docx) file using pure standard library
    zipfile + OpenXML generation, requiring zero pip packages.
    """
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    lines = content.splitlines()
    body_xml_parts = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Heading 1
        if stripped.startswith("# "):
            text = escape_xml(stripped[2:].strip())
            body_xml_parts.append(f"""
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/><w:spacing w:before="240" w:after="120"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="36"/><w:color w:val="1F497D"/></w:rPr><w:t>{text}</w:t></w:r>
    </w:p>""")
        # Heading 2
        elif stripped.startswith("## "):
            text = escape_xml(stripped[3:].strip())
            body_xml_parts.append(f"""
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/><w:spacing w:before="200" w:after="100"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="28"/><w:color w:val="2C5E8A"/></w:rPr><w:t>{text}</w:t></w:r>
    </w:p>""")
        # Heading 3
        elif stripped.startswith("### "):
            text = escape_xml(stripped[4:].strip())
            body_xml_parts.append(f"""
    <w:p>
      <w:pPr><w:pStyle w:val="Heading3"/><w:spacing w:before="160" w:after="80"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/><w:sz w:val="25"/><w:color w:val="3B6998"/></w:rPr><w:t>{text}</w:t></w:r>
    </w:p>""")
        # Blockquote
        elif stripped.startswith(">"):
            text = escape_xml(re.sub(r'^>\s?', '', stripped))
            body_xml_parts.append(f"""
    <w:p>
      <w:pPr><w:ind w:left="720"/><w:spacing w:before="100" w:after="120"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:i/><w:sz w:val="21"/><w:color w:val="555555"/></w:rPr><w:t>{text}</w:t></w:r>
    </w:p>""")
        # Regular paragraph with inline formatting
        else:
            tokens = parse_inline_markdown(stripped)
            runs_xml = []
            for t in tokens:
                t_escaped = escape_xml(t['text'])
                bold_tag = "<w:b/>" if t['bold'] else ""
                italic_tag = "<w:i/>" if t['italic'] else ""
                font_name = "Consolas" if t['code'] else "Calibri"
                color_val = "990000" if t['code'] else ("1B365D" if t['link'] else "222222")
                runs_xml.append(f"""<w:r><w:rPr><w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}"/>{bold_tag}{italic_tag}<w:sz w:val="22"/><w:color w:val="{color_val}"/></w:rPr><w:t xml:space="preserve">{t_escaped}</w:t></w:r>""")
            
            p_content = "".join(runs_xml)
            body_xml_parts.append(f"""
    <w:p>
      <w:pPr><w:spacing w:after="140" w:line="276" w:lineRule="auto"/></w:pPr>
      {p_content}
    </w:p>""")

    body_xml_inner = "\n".join(body_xml_parts)

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {body_xml_inner}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:color w:val="222222"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:rPr>
      <w:b/><w:sz w:val="36"/><w:color w:val="1F497D"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:rPr>
      <w:b/><w:sz w:val="28"/><w:color w:val="2C5E8A"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:rPr>
      <w:b/><w:sz w:val="25"/><w:color w:val="3B6998"/>
    </w:rPr>
  </w:style>
</w:styles>"""

    # Build ZIP archive (.docx)
    out_path = Path(docx_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(str(docx_path), "w", zipfile.ZIP_DEFLATED) as docx_zip:
        docx_zip.writestr("[Content_Types].xml", content_types_xml)
        docx_zip.writestr("_rels/.rels", rels_xml)
        docx_zip.writestr("word/document.xml", document_xml)
        docx_zip.writestr("word/styles.xml", styles_xml)


# ----------------------------------------------------------------------------
# 5. Core Conversion Dispatcher
# ----------------------------------------------------------------------------
def convert_file(input_file, output_file=None, font_family="Calibri"):
    """Converts a single markdown file to .docx."""
    in_p = Path(input_file).resolve()
    if not in_p.exists():
        raise FileNotFoundError(f"Input file not found: {in_p}")

    if output_file:
        out_p = Path(output_file).resolve()
    else:
        out_p = in_p.with_suffix(".docx")

    out_p.parent.mkdir(parents=True, exist_ok=True)

    print(f"[*] Converting: {in_p.name} -> {out_p.name}...")
    start_time = datetime.now()

    if HAVE_DOCX:
        convert_with_docx(str(in_p), str(out_p), font_family=font_family)
    else:
        convert_with_openxml_fallback(str(in_p), str(out_p))

    elapsed = (datetime.now() - start_time).total_seconds()
    file_size_kb = out_p.stat().st_size / 1024.0
    print(f"[+] Success! Generated '{out_p.name}' ({file_size_kb:.1f} KB in {elapsed:.2f}s)")
    return out_p


def convert_directory(input_dir, output_dir=None, recursive=False, font_family="Calibri"):
    """Batch converts all markdown files in a directory."""
    in_dir_p = Path(input_dir).resolve()
    if not in_dir_p.is_dir():
        raise NotADirectoryError(f"Directory not found: {in_dir_p}")

    pattern = "**/*.md" if recursive else "*.md"
    md_files = sorted(list(in_dir_p.glob(pattern)))

    if not md_files:
        print(f"[!] No markdown (.md) files found in: {in_dir_p}")
        return []

    print(f"[*] Found {len(md_files)} markdown file(s) in {in_dir_p}")
    results = []
    for md_file in md_files:
        if output_dir:
            out_dir_p = Path(output_dir).resolve()
            rel = md_file.relative_to(in_dir_p)
            out_f = (out_dir_p / rel).with_suffix(".docx")
        else:
            out_f = md_file.with_suffix(".docx")

        try:
            res = convert_file(md_file, out_f, font_family=font_family)
            results.append(res)
        except Exception as e:
            print(f"[!] Error converting {md_file.name}: {e}")

    print(f"[+] Finished batch conversion: {len(results)}/{len(md_files)} succeeded.")
    return results


# ----------------------------------------------------------------------------
# 6. Interactive Prompt & CLI Argument Parsing
# ----------------------------------------------------------------------------
def interactive_menu():
    """Provides a friendly interactive CLI when run with no arguments."""
    repo_root = Path(__file__).resolve().parent
    print("\n" + "=" * 65)
    print("      THEOS - MARKDOWN TO DOCX CONVERTER (LOCAL TOOL)")
    print("=" * 65)
    print("  [1] Convert a single Markdown file (.md)")
    print("  [2] Convert all .md files in a specific folder")
    print("  [3] Convert all .md files in the entire Theos workspace")
    print("  [0] Exit")
    print("=" * 65)

    choice = input("Enter choice (0-3): ").strip()

    if choice == "1":
        file_input = input("\nEnter path to .md file (or drag & drop here): ").strip().strip('"\'')
        if not file_input:
            print("No file specified.")
            return
        convert_file(file_input)
    elif choice == "2":
        dir_input = input("\nEnter folder path (or press Enter for current folder): ").strip().strip('"\'')
        if not dir_input:
            dir_input = str(repo_root)
        out_input = input("Enter output folder (or press Enter to save alongside .md files): ").strip().strip('"\'')
        out_dir = out_input if out_input else None
        convert_directory(dir_input, output_dir=out_dir, recursive=False)
    elif choice == "3":
        print("\nScanning entire workspace for all .md files...")
        convert_directory(str(repo_root), recursive=True)
    elif choice == "0":
        print("Goodbye!")
        return
    else:
        print("Invalid selection.")


def main():
    parser = argparse.ArgumentParser(
        description="Convert Markdown (.md) documents to Microsoft Word (.docx) files locally."
    )
    parser.add_argument("input", nargs="?", help="Path to a markdown file or directory to convert.")
    parser.add_argument("-o", "--output", help="Output .docx file path or directory.")
    parser.add_argument("--dir", help="Directory containing markdown files to convert.")
    parser.add_argument("--outdir", help="Directory where converted .docx files should be placed.")
    parser.add_argument("-r", "--recursive", action="store_true", help="Search subdirectories recursively.")
    parser.add_argument("--all", action="store_true", help="Convert all markdown files in current workspace.")
    parser.add_argument("--font", default="Calibri", help="Default font family (Calibri, Georgia, Aptos, etc.)")

    args = parser.parse_args()

    if args.all:
        repo_root = Path(__file__).resolve().parent
        convert_directory(str(repo_root), output_dir=args.outdir, recursive=True, font_family=args.font)
    elif args.dir:
        convert_directory(args.dir, output_dir=args.outdir, recursive=args.recursive, font_family=args.font)
    elif args.input:
        in_p = Path(args.input)
        if in_p.is_dir():
            convert_directory(str(in_p), output_dir=args.output, recursive=args.recursive, font_family=args.font)
        else:
            convert_file(str(in_p), output_file=args.output, font_family=args.font)
    else:
        interactive_menu()


if __name__ == "__main__":
    main()
