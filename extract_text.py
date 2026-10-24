import docx

doc_path = r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos\Cell-Leaders-Conference-D-ae2c0377-53f8-4bc8-9fd6-3a6d20381e80-2026-10-21-10-50-49.docx"
txt_path = r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos\raw_cell_leaders_transcript.txt"

doc = docx.Document(doc_path)
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
full_text = "\n\n".join(paragraphs)

with open(txt_path, "w", encoding="utf-8") as f:
    f.write(full_text)

print(f"Extracted {len(paragraphs)} paragraphs, {len(full_text.split())} words.")
