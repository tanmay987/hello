import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime, timezone

# ============================
# Namespace
# ============================
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ET.register_namespace("w", NS["w"])

def w_tag(tag: str) -> str:
    return f"{{{NS['w']}}}{tag}"

def now_w3c() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

# ============================
# updated.txt parsing
# ============================
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")
MARK_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)
WORD_RE = re.compile(r"\s+|[^\s]+")

def parse_updated_txt(path: str) -> dict[int, str]:
    out: dict[int, str] = {}
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line.strip():
                continue
            m = LINE_RE.match(line)
            if not m:
                continue
            out[int(m.group(1))] = m.group(2)
    return out

def split_text_tokens(text: str) -> list[str]:
    return [m.group(0) for m in WORD_RE.finditer(text)]

def tokenize_updated_atomic(s: str):
    tokens = []
    pos = 0
    for m in MARK_RE.finditer(s):
        outside = s[pos:m.start()]
        for t in split_text_tokens(outside):
            tokens.append(("TEXTTOK", t))

        kind = m.group(1)
        inner = m.group(2)
        tokens.append(("INSBLOCK" if kind == "INS" else "DELBLOCK", inner))

        pos = m.end()

    tail = s[pos:]
    for t in split_text_tokens(tail):
        tokens.append(("TEXTTOK", t))

    return tokens

# ============================
# Word text extraction
# ============================
def paragraph_plain_text(p: ET.Element) -> str:
    """Extract visible plain text in this paragraph (concatenate all w:t)."""
    parts = []
    for t in p.findall(".//w:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)

# ============================
# XML helpers
# ============================
def get_first_run_template(p: ET.Element) -> ET.Element:
    r = p.find(".//w:r", NS)
    return r if r is not None else ET.Element(w_tag("r"))

def make_plain_run(run_template: ET.Element, text: str) -> ET.Element:
    r = deepcopy(run_template)
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    t = ET.Element(w_tag("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r

def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str):
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)

    ins.append(make_plain_run(run_template, text))
    return ins

def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str):
    dele = ET.Element(w_tag("del"))
    dele.set(w_tag("id"), str(del_id))
    dele.set(w_tag("author"), author)
    dele.set(w_tag("date"), date)

    r = deepcopy(run_template)
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    dt = ET.Element(w_tag("delText"))
    dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    dt.text = text
    r.append(dt)

    dele.append(r)
    return dele

def clear_paragraph_runs(p: ET.Element):
    """Remove r/ins/del nodes, keep pPr intact."""
    for ch in list(p):
        if ch.tag in {w_tag("r"), w_tag("ins"), w_tag("del")}:
            p.remove(ch)

# ============================
# Strict rebuild paragraph
# ============================
def rebuild_paragraph_from_updated_strict_del(p: ET.Element, updated_marked: str, author: str, next_id: int):
    """
    Ground truth rebuild:
      - TEXTTOK and INS always applied
      - DEL only applied if del_text exists in original paragraph text
    """
    original_text = paragraph_plain_text(p)                 # original paragraph visible text
    tokens = tokenize_updated_atomic(updated_marked)
    date = now_w3c()
    template_run = get_first_run_template(p)

    clear_paragraph_runs(p)

    for kind, payload in tokens:
        if kind == "TEXTTOK":
            if payload:
                p.append(make_plain_run(template_run, payload))

        elif kind == "INSBLOCK":
            if payload:
                p.append(make_ins(template_run, payload, next_id, author, date))
                next_id += 1

        elif kind == "DELBLOCK":
            if not payload:
                continue

            # ✅ STRICT: only emit DEL if payload exists in original paragraph text
            if payload in original_text:
                p.append(make_del(template_run, payload, next_id, author, date))
                next_id += 1
            else:
                # ignore deletion block
                pass

    return next_id

# ============================
# Document-wide apply
# ============================
def apply_track_changes(document_xml_path: str, updated_txt_path: str, author: str = "Tanmay"):
    edits = parse_updated_txt(updated_txt_path)

    tree = ET.parse(document_xml_path)
    root = tree.getroot()
    paragraphs = root.findall(".//w:p", NS)

    next_id = 1
    changed = 0

    for i, p in enumerate(paragraphs, start=1):
        if i not in edits:
            continue
        next_id = rebuild_paragraph_from_updated_strict_del(p, edits[i], author, next_id)
        changed += 1

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Rebuilt {changed} paragraphs (strict DEL mode).")

if __name__ == "__main__":
    apply_track_changes("word/document.xml", "updated.txt", author="Tanmay")
