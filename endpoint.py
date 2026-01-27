from docx import Document
import io

# ========== 1) Extract text ==========

def extract_docx(file):
    """
    file: uploaded file (docx)
    returns: {"text": "..."}
    """
    content = file.read()
    doc = Document(io.BytesIO(content))

    text = "\n".join(p.text for p in doc.paragraphs)
    return {"text": text}


# ========== 2) Build docx ==========

def build_docx(text):
    """
    text: string
    returns: docx file
    """
    doc = Document()
    for line in text.split("\n"):
        doc.add_paragraph(line)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer
