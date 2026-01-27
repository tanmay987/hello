import dataiku
from dataiku.webapp import app
import zipfile
import io
import xml.etree.ElementTree as ET

@app.route("/")
def home():
    return "Backend started"

@app.route("/upload", methods=["POST"])
def upload():
    f = app.current_request.files.get("file")
    if f is None:
        return {"error": "no file"}

    data = f.read()

    with zipfile.ZipFile(io.BytesIO(data)) as z:
        xml_content = z.read("word/document.xml")

    root = ET.fromstring(xml_content)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    texts = []
    for node in root.findall(".//w:t", ns):
        if node.text:
            texts.append(node.text)

    return {"text": "\n".join(texts)}
