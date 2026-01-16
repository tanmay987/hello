import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from collections import Counter
from datetime import datetime, timezone

# ----------------------------
# Namespace
# ----------------------------
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ET.register_namespace("w", NS["w"])

def w_tag(tag: str) -> str:
    return f"{{{NS['w']}}}{tag}"

def now_w3c() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

# ----------------------------
# updated.txt parsing
# ----------------------------
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")
MARK_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)

# Tokenizer that preserves EXACT whitespace chunks
TOK_RE = re.compile(r"\s+|\S+")

def parse_updated_txt(path: str) -> dict[int, str]:
    out = {}
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

def strip_markers_keep_text(updated_marked: str) -> str:
    return MARK_RE.sub("", updated_marked)

def extract_all_del_text(updated_marked: str) -> list[str]:
    dels = []
    for m in MARK_RE.finditer(updated_marked):
        if m.group(1) == "DEL":
            dels.append(m.group(2))
    return dels

def tokenize_updated_atomic(updated_marked: str):
    """
    ("TEXT", chunk outside markers) chunks are kept as raw strings (no tokenization here),
    ("INS", ins_text) atomic,
    ("DEL", del_text) atomic.
    """
    tokens = []
    pos = 0
    for m in MARK_RE.finditer(updated_marked):
        outside = updated_marked[pos:m.start()]
        if outside:
            tokens.append(("TEXT", outside))

        kind = m.group(1)
        inner = m.group(2)
        tokens.append((kind, inner))

        pos = m.end()

    tail = updated_marked[pos:]
    if tail:
        tokens.append(("TEXT", tail))
    return tokens

def split_ws_tokens(s: str) -> list[str]:
    """Split into whitespace and non-whitespace tokens (exact whitespace preserved)."""
    return [m.group(0) for m in TOK_RE.finditer(s)]

def words_only(s: str) -> list[str]:
    """Return list of non-whitespace tokens."""
    return [t for t in split_ws_tokens(s) if not t.isspace()]

# ----------------------------
# Word XML utilities
# ----------------------------
def paragraph_runs(p: ET.Element):
    """Return list of run elements under paragraph."""
    return p.findall(".//w:r", NS)

def run_text(run: ET.Element) -> str:
    """Concatenate all w:t text nodes of a run."""
    parts = []
    for t in run.findall(".//w:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)

def set_run_text_preserve(run: ET.Element, text: str):
    """Replace run's w:t children with a single w:t containing text preserving spaces."""
    for node in list(run):
        if node.tag == w_tag("t"):
            run.remove(node)
    t = ET.Element(w_tag("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    run.append(t)

def build_paragraph_fulltext_and_runs(p: ET.Element):
    full_text = ""
    run_map = []
    for r in p.findall(".//w:r", NS):
        txt = run_text(r)
        if not txt:
            continue
        start = len(full_text)
        full_text += txt
        end = len(full_text)
        run_map.append({"r": r, "start": start, "end": end, "text": txt})
    return full_text, run_map

def find_run_covering_offset(run_map: list[dict], off: int):
    if not run_map:
        return None
    if off <= 0:
        return run_map[0]
    if off >= run_map[-1]["end"]:
        return run_map[-1]
    for m in run_map:
        if m["start"] <= off < m["end"]:
            return m
    return run_map[-1]

def split_run_at(run: ET.Element, offset: int):
    """Split a run at offset preserving formatting and preserving whitespace."""
    combined = run_text(run)

    left_text = combined[:offset]
    right_text = combined[offset:]

    left = deepcopy(run)
    right = deepcopy(run)

    set_run_text_preserve(left, left_text) if left_text != "" else set_run_text_preserve(left, "")
    set_run_text_preserve(right, right_text) if right_text != "" else set_run_text_preserve(right, "")

    # If empty text, remove w:t
    if left_text == "":
        for node in list(left):
            if node.tag == w_tag("t"):
                left.remove(node)
    if right_text == "":
        for node in list(right):
            if node.tag == w_tag("t"):
                right.remove(node)

    return left, right

def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str):
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)

    r = deepcopy(run_template)
    set_run_text_preserve(r, text)
    ins.append(r)
    return ins

def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str):
    dele = ET.Element(w_tag("del"))
    dele.set(w_tag("id"), str(del_id))
    dele.set(w_tag("author"), author)
    dele.set(w_tag("date"), date)

    r = deepcopy(run_template)

    # remove w:t
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    dt = ET.Element(w_tag("delText"))
    dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    dt.text = text
    r.append(dt)

    dele.append(r)
    return dele

# ----------------------------
# Step 1: filter doc by updated words, preserving all whitespace
# ----------------------------
def filter_paragraph_preserve_spaces(p: ET.Element, keep_words_counter: Counter):
    """
    Remove ONLY word-tokens that are not present in keep_words_counter.
    Preserve whitespace tokens exactly as in the document.
    """
    for r in paragraph_runs(p):
        txt = run_text(r)
        if not txt:
            continue

        toks = split_ws_tokens(txt)
        out = []

        for tok in toks:
            if tok.isspace():
                # ✅ always keep whitespace exactly
                out.append(tok)
                continue

            # word token
            if keep_words_counter[tok] > 0:
                out.append(tok)
                keep_words_counter[tok] -= 1
            else:
                # drop word token but keep surrounding whitespace
                pass

        new_txt = "".join(out)

        # Update run text (keep run formatting)
        set_run_text_preserve(r, new_txt)

# ----------------------------
# Step 2: apply DEL by word-sequence match + INS by next-anchor
# ----------------------------
def find_word_sequence_span(full_text: str, del_words: list[str], start_char: int = 0):
    """
    Find character span (start,end) in full_text where del_words appear in sequence,
    allowing arbitrary whitespace between them.
    """
    if not del_words:
        return None

    # Build regex like: r'\bin\b\s+\bcar\b'
    # But since tokens may include punctuation, use escaping + \s+
    pattern = r"\s+".join(re.escape(w) for w in del_words)
    rgx = re.compile(pattern)

    m = rgx.search(full_text, start_char)
    if not m:
        return None
    return (m.start(), m.end())

def apply_delete_span(old_children, run_map, new_children, child_idx_ref, start_off, end_off, del_text, author, next_id, date):
    """
    Remove original normal text [start_off, end_off) and replace it with one <w:del>.
    """
    first_rm = find_run_covering_offset(run_map, start_off)
    last_rm = find_run_covering_offset(run_map, max(start_off, end_off - 1))
    if not first_rm or not last_rm:
        return next_id

    first_run = first_rm["r"]
    last_run = last_rm["r"]

    # copy children until first_run
    while child_idx_ref[0] < len(old_children):
        ch = old_children[child_idx_ref[0]]
        if ch is first_run:
            break
        new_children.append(ch)
        child_idx_ref[0] += 1

    if child_idx_ref[0] >= len(old_children) or old_children[child_idx_ref[0]] is not first_run:
        new_children.append(make_del(first_run, del_text, next_id, author, date))
        return next_id + 1

    first_local = start_off - first_rm["start"]
    last_local = end_off - last_rm["start"]

    left_first, right_first = split_run_at(first_run, first_local)

    if first_run is last_run:
        _mid, right_last = split_run_at(right_first, last_local - first_local)
    else:
        _mid, right_last = split_run_at(last_run, last_local)

    child_idx_ref[0] += 1

    if first_run is not last_run:
        while child_idx_ref[0] < len(old_children):
            ch = old_children[child_idx_ref[0]]
            child_idx_ref[0] += 1
            if ch is last_run:
                break

    if left_first.findall(".//w:t", NS):
        new_children.append(left_first)

    new_children.append(make_del(first_run, del_text, next_id, author, date))
    next_id += 1

    if right_last.findall(".//w:t", NS):
        new_children.append(right_last)

    return next_id

def apply_markers_after_filter(p: ET.Element, updated_marked: str, author: str, next_id: int):
    full_text, run_map = build_paragraph_fulltext_and_runs(p)
    if not full_text or not run_map:
        return next_id

    tokens = tokenize_updated_atomic(updated_marked)
    date = now_w3c()

    cursor_char = 0
    events = []

    # Build “future anchors” from TEXT chunks (raw)
    # For insertion: next non-empty word token after INS
    chunks = []
    for k, v in tokens:
        chunks.append((k, v))

    for idx, (kind, payload) in enumerate(chunks):
        if kind == "DEL":
            del_text = payload
            if not del_text.strip():
                continue

            del_words = words_only(del_text)
            span = find_word_sequence_span(full_text, del_words, start_char=cursor_char)
            if span:
                s, e = span
                events.append((s, "DEL", (e, del_text)))
                cursor_char = e
            continue

        if kind == "INS":
            ins_text = payload
            if not ins_text:
                continue

            # next anchor word
            anchor_word = None
            for j in range(idx + 1, len(chunks)):
                if chunks[j][0] == "TEXT":
                    for w in words_only(chunks[j][1]):
                        anchor_word = w
                        break
                if anchor_word:
                    break

            if anchor_word:
                pos = full_text.find(anchor_word, cursor_char)
                insert_pos = pos if pos >= 0 else cursor_char
            else:
                insert_pos = len(full_text)

            events.append((insert_pos, "INS", ins_text))
            continue

    if not events:
        return next_id

    # DEL before INS at same offset
    events.sort(key=lambda x: (x[0], 0 if x[1] == "DEL" else 1))

    old_children = list(p)
    new_children = []
    child_idx_ref = [0]

    for off, kind, payload in events:
        if kind == "DEL":
            end_off, del_text = payload
            next_id = apply_delete_span(
                old_children, run_map, new_children, child_idx_ref,
                off, end_off, del_text, author, next_id, date
            )
        else:
            ins_text = payload
            rm = find_run_covering_offset(run_map, off)
            base_run = rm["r"] if rm else run_map[0]["r"]

            while child_idx_ref[0] < len(old_children):
                ch = old_children[child_idx_ref[0]]
                if ch is base_run:
                    break
                new_children.append(ch)
                child_idx_ref[0] += 1

            if child_idx_ref[0] >= len(old_children) or old_children[child_idx_ref[0]] is not base_run:
                new_children.append(make_ins(base_run, ins_text, next_id, author, date))
                next_id += 1
                continue

            local = off - rm["start"]
            local = max(0, min(local, rm["end"] - rm["start"]))

            left_run, right_run = split_run_at(base_run, local)

            child_idx_ref[0] += 1

            if left_run.findall(".//w:t", NS):
                new_children.append(left_run)

            new_children.append(make_ins(base_run, ins_text, next_id, author, date))
            next_id += 1

            if right_run.findall(".//w:t", NS):
                new_children.append(right_run)

    while child_idx_ref[0] < len(old_children):
        new_children.append(old_children[child_idx_ref[0]])
        child_idx_ref[0] += 1

    for ch in list(p):
        p.remove(ch)
    for ch in new_children:
        p.append(ch)

    return next_id

# ----------------------------
# Main apply
# ----------------------------
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

        updated_marked = edits[i]

        # Keep words from outside text
        updated_keep_text = strip_markers_keep_text(updated_marked)
        keep_words = words_only(updated_keep_text)

        # Also keep words inside DEL (so they remain to be wrapped)
        del_words = []
        for d in extract_all_del_text(updated_marked):
            del_words.extend(words_only(d))

        keep_counter = Counter(keep_words + del_words)

        # Step 1: filter doc by keep words (spaces preserved)
        filter_paragraph_preserve_spaces(p, keep_counter)

        # Step 2: apply INS/DEL markers
        next_id = apply_markers_after_filter(p, updated_marked, author, next_id)

        changed += 1

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Updated {changed} paragraphs (doc truth, spaces preserved, DEL works).")

if __name__ == "__main__":
    apply_track_changes("word/document.xml", "updated.txt", author="Tanmay")
