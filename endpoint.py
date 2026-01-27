import zipfile
import io
import xml.etree.ElementTree as ET
from flask import Flask, request, jsonify

app = Flask(__name__)

# ------------------------------------------------
# Extract text from word/document.xml inside DOCX
# ------------------------------------------------

def extract_text_from_docx_bytes(docx_bytes):
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        xml_content = z.read("word/document.xml")

    root = ET.fromstring(xml_content)

    # Word namespace
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    texts = []
    for node in root.findall(".//w:t", ns):
        if node.text:
            texts.append(node.text)

    return "\n".join(texts)


# ------------------------------------------------
# Upload endpoint
# ------------------------------------------------

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    data = f.read()

    text = extract_text_from_docx_bytes(data)

    return jsonify({"text": text})
