from imghdr import tests

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

def process_paragraphs(doc, start_keyword, occurrence, stop_keyword=None):
    collecting = False
    current_paragraph = []
    paragraphs_between = []
    counter = 0

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()

        if not text:
            continue

        if start_keyword.lower() in text.lower():
            counter+=1
            if counter == occurrence:
                collecting = True
                text = paragraph.text[paragraph.text.lower().find(start_keyword):]
                parts = text.split("“")
                if len(parts) > 1:
                    before_quote = parts[0].strip()  # Text before the quote
                    quoted_text, after_quote = parts[1].split("”")  # Quoted text and remaining part

                # Clean up the remaining part
                after_quote = after_quote.strip()

                # Combine the parts with newline characters
                result = f"{before_quote}\n“{quoted_text}”\n{after_quote}"

                current_paragraph.append(result)
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

    if collecting and current_paragraph:
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

def singing(prs, docx_path, occurrence, delimiter, language):
    doc = Document(docx_path)
    start_keyword = "marende" if language == 1 else "bernyanyi"
    paragraphs = process_paragraphs(doc, start_keyword, occurrence, delimiter)
    font_size_map = {250: 36, 200: 40, 0: 48}
    add_paragraphs_to_slide(prs, paragraphs, font_size_map)

def patik(prs, docx_path, language):
    doc = Document(docx_path)
    collecting = False
    paragraphs = []
    paragraphs_between = []
    paragraph_count = 0
    delimiter = "marende" if language == 1 else "bernyanyi"

    for paragraph in doc.paragraphs:
        paragraph_count += 1
        if "p a t i k".lower() in paragraph.text.lower():
            collecting = True
            paragraphs.append("P A T I K")
            text = paragraph.text[paragraph.text.lower().find("p a t i k".lower()):]
            text = text.replace("P a t i k :", "").strip()
            if text:
                paragraphs.append(text)

            continue

        if collecting:
            if delimiter in paragraph.text.lower():
                break

            paragraphs.append(paragraph.text)

    if collecting and paragraphs:
        paragraphs_between.append("\n".join(paragraphs))

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
    cleaned_texts = []
    for text in found_text:
        text = text.lstrip(''.join(c for c in text if not c.isalpha() and not c.isspace())).rstrip(':').upper()
        split_text = text.split(":", 1)
        if len(split_text) > 1:
            modified_text = split_text[0] + ":\n" + split_text[1]
            cleaned_texts.append(modified_text)
        else:
            cleaned_texts.append(text)
    font_size_map = {250: 36, 200: 40, 0: 48}
    add_paragraphs_to_slide(prs, cleaned_texts, font_size_map)

def epistel(prs, docx_path, language):
    doc = Document(docx_path)
    paragraphs = []  # List to store the final grouped paragraphs
    current_paragraph = []  # Temporarily stores lines for a single speaker
    found_epistel = False
    found_marende = False

    marende = "marende" if language == 1 else "bernyanyi"

    # Process each paragraph in the document
    for para in doc.paragraphs:
        text = para.text.strip()  # Clean up extra whitespace
        if not text:  # Skip empty lines
            continue

        if "E P I S T E L".lower() in text.lower():
            found_epistel = True
            continue

        if found_epistel and marende.lower() in text.lower():
            found_marende = True
            continue

        if found_epistel and not found_marende:
            # Check if the line starts with "U:" or "H:"
            if text.startswith("U	:") or text.startswith("H	:") or text.startswith("H :") or text.startswith("U :") or text.startswith("P	:") or text.startswith("J	:")or text.startswith("P :") or text.startswith("J :"):
                # Save the current paragraph if it exists
                if current_paragraph:
                    paragraphs.append("\n".join(current_paragraph))
                    current_paragraph = []  # Reset for the next speaker

                # Start a new paragraph
                current_speaker = text[:2]  # Extract the speaker identifier
                current_paragraph.append(text)
            else:
                # Add the line to the current paragraph
                current_paragraph.append(text)

    # Append the last paragraph if any
    if current_paragraph:
        paragraphs.append("\n".join(current_paragraph))

    font_size_map = {250: 36, 200: 40, 0: 48}
    add_paragraphs_to_slide(prs, paragraphs, font_size_map)


def convert_with_cover(docx_path, pptx_path, config, language):
    prs = Presentation()

    # Generate cover slide
    generate_cover(prs, docx_path)

    # Loop through the configuration and perform the steps dynamically
    for step in config['steps']:
        if step['action'] == 'singing':
            singing(prs, docx_path, step['number'], step['label'], language)
        elif step['action'] == 'session':
            session(prs, docx_path, step['label'])
        elif step['action'] == 'epistel':
            epistel(prs, docx_path, language)
        elif step['action'] == 'patik':
            patik(prs, docx_path, language)

    try:
        prs.save(pptx_path)
    except PermissionError:
        raise Exception(f"Cannot save to {pptx_path}. File may be open or you don't have write permissions.")
    except OSError as e:
        raise Exception(f"Error saving presentation to {pptx_path}: {str(e)}")
