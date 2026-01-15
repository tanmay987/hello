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

# TEXT matching should be robust, but WORD tokenization is enough
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
    """Tokenize TEXT outside markers into word/space tokens (for alignment only)."""
    return [m.group(0) for m in WORD_RE.finditer(text)]

def tokenize_updated_atomic(s: str):
    """
    Tokenize updated paragraph into:
      ("TEXTTOK", token)  -> word/space token OUTSIDE markers
      ("INSBLOCK", text)  -> entire raw INS block content
      ("DELBLOCK", text)  -> entire raw DEL block content

    INS/DEL are atomic (not tokenized).
    """
    tokens = []
    pos = 0
    for m in MARK_RE.finditer(s):
        outside = s[pos:m.start()]
        for t in split_text_tokens(outside):
            tokens.append(("TEXTTOK", t))

        kind = m.group(1)
        inner = m.group(2)  # whole text inside marker
        if kind == "INS":
            tokens.append(("INSBLOCK", inner))
        else:
            tokens.append(("DELBLOCK", inner))

        pos = m.end()

    tail = s[pos:]
    for t in split_text_tokens(tail):
        tokens.append(("TEXTTOK", t))

    return tokens

# ============================
# document.xml paragraph run map
# ============================
def build_paragraph_fulltext_and_runs(p: ET.Element):
    full_text = ""
    runs = []
    for r in p.findall(".//w:r", NS):
        texts = []
        for t in r.findall(".//w:t", NS):
            if t.text:
                texts.append(t.text)
        run_text = "".join(texts)
        if run_text:
            start = len(full_text)
            full_text += run_text
            end = len(full_text)
            runs.append({"r": r, "text": run_text, "start": start, "end": end})
    return full_text, runs

def find_run_at_offset(run_map: list[dict], char_off: int):
    if not run_map:
        return None
    if char_off <= 0:
        return run_map[0]
    if char_off >= run_map[-1]["end"]:
        return run_map[-1]
    for m in run_map:
        if m["start"] <= char_off < m["end"]:
            return m
    return run_map[-1]

def split_run_at(run: ET.Element, offset: int):
    """
    Split run at offset and ALWAYS preserve spaces.
    """
    t_nodes = run.findall(".//w:t", NS)
    combined = "".join([(t.text or "") for t in t_nodes])

    left_text = combined[:offset]
    right_text = combined[offset:]

    left = deepcopy(run)
    right = deepcopy(run)

    for node in list(left):
        if node.tag == w_tag("t"):
            left.remove(node)
    for node in list(right):
        if node.tag == w_tag("t"):
            right.remove(node)

    def add_wt(run_elem: ET.Element, txt: str):
        if txt == "":
            return
        t = ET.Element(w_tag("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = txt
        run_elem.append(t)

    add_wt(left, left_text)
    add_wt(right, right_text)
    return left, right

# ============================
# Track change nodes (atomic text)
# ============================
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

def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str):
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)

    r = deepcopy(run_template)
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    t = ET.Element(w_tag("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)

    ins.append(r)
    return ins

# ============================
# Apply markers with full-block INS/DEL
# ============================
def apply_markers_atomic(p: ET.Element, updated_marked: str, author: str, next_id: int, del_lookahead: int = 50):
    """
    Applies:
      - INS as ONE block insertion
      - DEL as ONE block deletion
      - ignores differences outside markers
    """
    original_text, run_map = build_paragraph_fulltext_and_runs(p)
    if not original_text or not run_map:
        return next_id

    # document tokens for alignment (TEXTTOK only)
    doc_tokens = split_text_tokens(original_text)

    # map token index -> char offset
    doc_tok_char = []
    cur = 0
    for tok in doc_tokens:
        doc_tok_char.append(cur)
        cur += len(tok)

    upd = tokenize_updated_atomic(updated_marked)

    doc_i = 0
    cursor_char = 0

    # Events: (char_offset, kind, text)
    events = []

    for kind, payload in upd:
        if kind == "INSBLOCK":
            if payload:
                events.append((cursor_char, "INS", payload))
            continue

        if kind == "DELBLOCK":
            # delete entire payload at best match location
            del_text = payload
            if not del_text:
                continue

            # Attempt alignment token-wise: tokenize deletion payload like doc tokens
            del_tokens = split_text_tokens(del_text)

            # Find where del_tokens appear in doc_tokens starting at doc_i
            found = -1
            max_j = min(len(doc_tokens), doc_i + del_lookahead)
            for j in range(doc_i, max_j):
                if doc_tokens[j:j + len(del_tokens)] == del_tokens:
                    found = j
                    break

            if found >= 0:
                # advance doc cursor to found
                while doc_i < found:
                    cursor_char += len(doc_tokens[doc_i])
                    doc_i += 1

                # delete block at cursor_char
                events.append((cursor_char, "DEL", del_text))

                # consume those tokens from doc stream
                for _ in range(len(del_tokens)):
                    cursor_char += len(doc_tokens[doc_i])
                    doc_i += 1
            else:
                # cannot align => ignore deletion
                continue
            continue

        # TEXTTOK alignment only
        tok = payload
        if doc_i < len(doc_tokens) and doc_tokens[doc_i] == tok:
            cursor_char += len(tok)
            doc_i += 1
        else:
            # ignore mismatch outside markers
            continue

    if not events:
        return next_id

    # Sort by offset, DEL before INS at same offset
    events.sort(key=lambda x: (x[0], 0 if x[1] == "DEL" else 1))

    old_children = list(p)
    new_children = []
    child_idx = 0
    date = now_w3c()

    consumed_runs = set()

    def copy_children_until(target_run: ET.Element):
        nonlocal child_idx
        while child_idx < len(old_children):
            ch = old_children[child_idx]
            if ch is target_run:
                break
            new_children.append(ch)
            child_idx += 1

    for off, kind, text in events:
        rm = find_run_at_offset(run_map, off)
        if rm is None:
            continue
        base_run = rm["r"]

        if base_run in consumed_runs:
            # find a nearby unused run
            candidate = None
            for m in run_map:
                if m["r"] not in consumed_runs and m["start"] <= off < m["end"]:
                    candidate = m
                    break
            if candidate is None:
                for m in run_map:
                    if m["r"] not in consumed_runs and m["start"] >= off:
                        candidate = m
                        break
            rm = candidate if candidate else run_map[-1]
            base_run = rm["r"]

        copy_children_until(base_run)

        # if we cannot locate precisely => append
        if child_idx >= len(old_children) or old_children[child_idx] is not base_run:
            template = rm["r"]
            if kind == "INS":
                new_children.append(make_ins(template, text, next_id, author, date))
                next_id += 1
            else:
                new_children.append(make_del(template, text, next_id, author, date))
                next_id += 1
            continue

        local = off - rm["start"]
        local = max(0, min(local, rm["end"] - rm["start"]))

        left_run, right_run = split_run_at(base_run, local)

        # consume base run
        child_idx += 1
        consumed_runs.add(base_run)

        if left_run.findall(".//w:t", NS):
            new_children.append(left_run)

        if kind == "INS":
            new_children.append(make_ins(base_run, text, next_id, author, date))
            next_id += 1
        else:
            new_children.append(make_del(base_run, text, next_id, author, date))
            next_id += 1

        if right_run.findall(".//w:t", NS):
            new_children.append(right_run)

    # add remaining original children
    while child_idx < len(old_children):
        new_children.append(old_children[child_idx])
        child_idx += 1

    # replace paragraph children
    for ch in list(p):
        p.remove(ch)
    for ch in new_children:
        p.append(ch)

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

        marked = edits[i]
        if "<INS:" not in marked and "<DEL:" not in marked:
            continue

        next_id = apply_markers_atomic(p, marked, author, next_id)
        changed += 1

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Applied atomic marker tracked changes to {changed} paragraphs.")


if __name__ == "__main__":
    apply_track_changes("word/document.xml", "updated.txt", author="Tanmay")
