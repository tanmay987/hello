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
    """
    Token stream:
      ("TEXTTOK", token) -> word/space tokens outside markers
      ("INSBLOCK", text) -> whole INS content
      ("DELBLOCK", text) -> whole DEL content
    """
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
# Paragraph extraction
# ============================
def extract_run_spans(p: ET.Element):
    """
    Extract original paragraph as a list of spans:
      (start_off, end_off, run_template, text)

    Where run_template is the original <w:r> element.

    We use .//w:r to include hyperlink runs etc.
    """
    spans = []
    off = 0

    for r in p.findall(".//w:r", NS):
        texts = []
        for t in r.findall(".//w:t", NS):
            if t.text:
                texts.append(t.text)
        run_text = "".join(texts)

        if not run_text:
            continue

        start = off
        end = off + len(run_text)
        spans.append((start, end, r, run_text))
        off = end

    full_text = "".join([s[3] for s in spans])
    return full_text, spans

def paragraph_plain_text(p: ET.Element) -> str:
    parts = []
    for t in p.findall(".//w:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)

# ============================
# Formatting mapping
# ============================
def pick_template_for_offset(spans, char_off: int):
    """
    Pick a run template based on where char_off falls in original spans.
    If out of range, return nearest run.
    """
    if not spans:
        return ET.Element(w_tag("r"))

    if char_off <= spans[0][0]:
        return spans[0][2]
    if char_off >= spans[-1][1]:
        return spans[-1][2]

    for start, end, run, _txt in spans:
        if start <= char_off < end:
            return run

    return spans[-1][2]

# ============================
# XML builders
# ============================
def clear_paragraph_runs(p: ET.Element):
    """Remove all runs/ins/del, keep pPr intact."""
    for ch in list(p):
        if ch.tag in {w_tag("r"), w_tag("ins"), w_tag("del")}:
            p.remove(ch)

def make_plain_run(run_template: ET.Element, text: str) -> ET.Element:
    """
    Create normal run with same formatting as run_template.
    Always preserve spaces.
    """
    r = deepcopy(run_template)
    # remove direct <w:t> children
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

# ============================
# Core: rebuild paragraph with formatting + strict del
# ============================
def rebuild_paragraph_truth_preserve_format_strict_del(
    p: ET.Element,
    updated_marked: str,
    author: str,
    next_id: int,
):
    """
    Ground-truth rebuild:
      - updated_marked defines content
      - TEXTTOK: normal run
      - INS: tracked insertion
      - DEL: tracked deletion ONLY if payload exists in original paragraph text
      - formatting: each emitted segment uses run template chosen by original char offset
    """
    original_text, spans = extract_run_spans(p)
    strict_original_plain = original_text  # visible original text

    tokens = tokenize_updated_atomic(updated_marked)
    date = now_w3c()

    clear_paragraph_runs(p)

    # We'll track an approximate "cursor" into original_text for formatting selection
    # TEXTTOK advances cursor when it matches at cursor; otherwise cursor stays.
    cursor = 0

    for kind, payload in tokens:
        if kind == "TEXTTOK":
            if payload == "":
                continue

            # pick formatting near current cursor
            template = pick_template_for_offset(spans, cursor)

            # emit normal run
            p.append(make_plain_run(template, payload))

            # advance cursor if it matches original at this position (best-effort)
            if strict_original_plain.startswith(payload, cursor):
                cursor += len(payload)
            else:
                # mismatch => do NOT move cursor (updated is truth)
                pass

        elif kind == "INSBLOCK":
            if payload == "":
                continue

            # insertion: use formatting at cursor position
            template = pick_template_for_offset(spans, cursor)
            p.append(make_ins(template, payload, next_id, author, date))
            next_id += 1

            # insertion isn't in original => do not move cursor

        elif kind == "DELBLOCK":
            if payload == "":
                continue

            # STRICT deletion: only emit if payload exists somewhere in original paragraph
            if payload in strict_original_plain:
                # choose formatting where it occurs (closest after cursor if possible)
                pos = strict_original_plain.find(payload, cursor)
                if pos < 0:
                    pos = strict_original_plain.find(payload)

                template = pick_template_for_offset(spans, pos if pos >= 0 else cursor)
                p.append(make_del(template, payload, next_id, author, date))
                next_id += 1

                # if deleted text occurs after cursor, advance cursor past it
                if pos >= cursor and pos >= 0:
                    cursor = pos + len(payload)
            else:
                # ignore deletion if not present in original
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
        next_id = rebuild_paragraph_truth_preserve_format_strict_del(
            p=p,
            updated_marked=edits[i],
            author=author,
            next_id=next_id,
        )
        changed += 1

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Rebuilt {changed} paragraphs (updated truth + strict DEL + preserved formatting).")

if __name__ == "__main__":
    apply_track_changes("word/document.xml", "updated.txt", author="Tanmay")
