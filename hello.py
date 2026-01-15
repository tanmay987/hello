import xml.etree.ElementTree as ET

# WordprocessingML namespace
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def extract_plain(document_xml_path: str, out_txt: str):
    tree = ET.parse(document_xml_path)                          # load document.xml into memory
    root = tree.getroot()                                       # get root element <w:document>

    paragraphs = root.findall(".//w:p", NS)                     # find all paragraph nodes in the document

    with open(out_txt, "w", encoding="utf-8") as f:             # open output file
        for i, p in enumerate(paragraphs, start=1):             # enumerate paragraphs (1-based)
            parts = []                                          # collect plain text fragments
            for t in p.findall(".//w:t", NS):                   # find all <w:t> nodes (visible text)
                if t.text:                                      # if node actually has text
                    parts.append(t.text)                        # append fragment
            para_text = "".join(parts).replace("\n", " ")       # join fragments into paragraph text
            f.write(f"[P{i}] {para_text}\n")                    # write paragraph line

    print(f"✅ Extracted {len(paragraphs)} paragraphs to {out_txt}")

# ---- Usage ----
# extract_plain("word/document.xml", "plain.txt")

if __name__ == "__main__":
    extract_plain("word/document.xml", "plain.txt")
