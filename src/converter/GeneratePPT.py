from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import re

def add_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    for shape in slide.shapes:
        if shape.is_placeholder:
            sp = shape.element
            sp.getparent().remove(sp)
    text_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    return text_frame

def process_paragraphs(doc, start_keyword, stop_keyword=None):
    collecting = False
    current_paragraph = []
    paragraphs_between = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        if start_keyword.lower() in text.lower():
            collecting = True
            current_paragraph.append(text)
            continue

        if collecting:
            if stop_keyword and stop_keyword.lower() in text.lower():
                break

            if "♫♪♫" in text or "musik" in text.lower():
                if current_paragraph:
                    paragraphs_between.append("\n".join(current_paragraph))
                    current_paragraph = []
                cleaned_text = text.replace("♫♪♫", "").strip()
                current_paragraph.append("MUSIK" if "musik" in cleaned_text.lower() else cleaned_text)
            else:
                current_paragraph.append(text)

    if current_paragraph:
        paragraphs_between.append("\n".join(current_paragraph))

    return paragraphs_between

def add_paragraphs_to_slide(prs, paragraphs, font_size_map):
    for paragraph in paragraphs:
        text_frame = add_slide(prs)
        p = text_frame.paragraphs[0]
        p.text = paragraph
        p.alignment = PP_ALIGN.CENTER
        paragraph_length = len(paragraph)

        for length, size in font_size_map.items():
            if paragraph_length > length:
                p.font.size = Pt(size)
                break
        else:
            p.font.size = Pt(font_size_map[0])

def singing(prs, docx_path, occurance, delimiter):
    doc = Document(docx_path)
    paragraphs = process_paragraphs(doc, "marende", delimiter)
    font_size_map = {250: 36, 200: 40, 0: 48}
    add_paragraphs_to_slide(prs, paragraphs, font_size_map)

def patik(prs, docx_path):
    doc = Document(docx_path)
    paragraphs = process_paragraphs(doc, "p a t i k", "marende")
    font_size_map = {230: 14, 200: 36, 0: 48}
    add_paragraphs_to_slide(prs, paragraphs, font_size_map)

def generate_cover(prs, docx_path):
    doc = Document(docx_path)
    page_text = []
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():
            continue

        has_page_break = any("w:br" in run._element.xml and 'type="page"' in run._element.xml for run in paragraph.runs)
        if has_page_break:
            break

        page_text.append(paragraph.text.strip())

    text_frame = add_slide(prs)
    for line in "\n".join(page_text).split("\n"):
        p = text_frame.add_paragraph()
        p.text = line
        p.alignment = PP_ALIGN.CENTER

        font_size = 36
        if re.search(r"topik|huria", line, re.IGNORECASE):
            p.text = "\n" + p.text
        if "partording" in line.lower():
            font_size = 48
        elif any(keyword in line.lower() for keyword in ["ev", "ep", ",", "huria"]):
            font_size = 24

        p.font.size = Pt(font_size)

def session(prs, docx_path, param):
    doc = Document(docx_path)
    found_text = [paragraph.text.strip() for paragraph in doc.paragraphs if param.lower() in paragraph.text.lower()]
    cleaned_texts = [
        text.lstrip(''.join(c for c in text if not c.isalpha() and not c.isspace())).rstrip(':').upper()
        for text in found_text
    ]
    font_size_map = {250: 36, 200: 40, 0: 48}
    add_paragraphs_to_slide(prs, cleaned_texts, font_size_map)

def epistel(prs, docx_path):
    doc = Document(docx_path)
    paragraphs = process_paragraphs(doc, "e p i s t e l", "marende")
    font_size_map = {250: 36, 200: 40, 0: 48}
    add_paragraphs_to_slide(prs, paragraphs, font_size_map)

def convert_with_cover(docx_path, pptx_path):
    prs = Presentation()
    generate_cover(prs, docx_path)
    singing(prs, docx_path, 1, "votum")
    session(prs, docx_path, "votum")
    singing(prs, docx_path, 2, "p a t i k")
    patik(prs, docx_path)
    singing(prs, docx_path, 3, "manopoti dosa")
    session(prs, docx_path, "manopoti dosa")
    singing(prs, docx_path, 4, "e p i s t e l")
    session(prs, docx_path, "e p i s t e l")
    epistel(prs, docx_path)
    singing(prs, docx_path, 5, "manghatindanghon haporseaon")
    session(prs, docx_path, "manghatindanghon haporseaon")
    session(prs, docx_path, "koor")
    session(prs, docx_path, "tingting")
    session(prs, docx_path, "sunggul")
    prs.save(pptx_path)
