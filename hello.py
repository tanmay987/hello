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
    """Return fully-qualified WordprocessingML tag."""
    return f"{{{NS['w']}}}{tag}"

def now_w3c() -> str:
    """UTC timestamp in Word track-changes format."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

# ----------------------------
# updated.txt parsing
# ----------------------------
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")
MARK_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)
WORD_RE = re.compile(r"\s+|[^\s]+")

def parse_updated_txt(path: str) -> dict[int, str]:
    """Parse updated.txt into {paragraph_id: marked_content}."""
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

def split_tokens(text: str) -> list[str]:
    """Split text into word/space tokens (spaces preserved)."""
    return [m.group(0) for m in WORD_RE.finditer(text)]

def strip_markers_keep_text(updated_marked: str) -> str:
    """Remove <INS:..> and <DEL:..> blocks, leaving only outside text."""
    return MARK_RE.sub("", updated_marked)

def extract_all_del_text(updated_marked: str) -> list[str]:
    """Return list of all deletion block contents (<DEL:...>) in updated_marked."""
    dels = []
    for m in MARK_RE.finditer(updated_marked):
        if m.group(1) == "DEL":
            dels.append(m.group(2))
    return dels

def tokenize_updated_atomic(updated_marked: str):
    """
    Token stream:
      ("TEXTTOK", token) -> outside markers (tokenized word/space)
      ("INSBLOCK", text) -> atomic INS
      ("DELBLOCK", text) -> atomic DEL
    """
    tokens = []
    pos = 0
    for m in MARK_RE.finditer(updated_marked):
        outside = updated_marked[pos:m.start()]
        for t in split_tokens(outside):
            tokens.append(("TEXTTOK", t))
        kind = m.group(1)
        inner = m.group(2)
        tokens.append(("INSBLOCK" if kind == "INS" else "DELBLOCK", inner))
        pos = m.end()
    tail = updated_marked[pos:]
    for t in split_tokens(tail):
        tokens.append(("TEXTTOK", t))
    return tokens

# ----------------------------
# Word XML utilities
# ----------------------------
def build_paragraph_fulltext_and_runs(p: ET.Element):
    """Return (full_text, run_map) for ALL runs in the paragraph."""
    full_text = ""
    run_map = []
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
            run_map.append({"r": r, "start": start, "end": end, "text": run_text})
    return full_text, run_map

def find_run_covering_offset(run_map: list[dict], off: int):
    """Find run-map entry covering character offset 'off'."""
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
    """
    Split a run into left/right at offset preserving formatting.
    ALWAYS preserve spaces to avoid word-collisions.
    """
    t_nodes = run.findall(".//w:t", NS)
    combined = "".join([(t.text or "") for t in t_nodes])

    left_text = combined[:offset]
    right_text = combined[offset:]

    left = deepcopy(run)
    right = deepcopy(run)

    # Remove direct w:t children from clones
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

def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str):
    """Create a single <w:ins> containing the entire inserted text."""
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

def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str):
    """Create a single <w:del> containing the entire deleted text."""
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

# ----------------------------
# Step 1: filter doc runs using updated outside-text + DEL text
# ----------------------------
def filter_paragraph_by_updated_text(p: ET.Element, updated_keep_counter: Counter):
    """
    Remove tokens from document.xml paragraph that do not appear in updated_keep_counter.

    IMPORTANT:
      - Keep tokens that appear in updated outside markers
      - ALSO keep tokens inside <DEL:...> blocks (so they can be wrapped later)
    """
    new_children = []

    for ch in list(p):
        # Keep non-run children unchanged (e.g., pPr)
        if ch.tag != w_tag("r"):
            new_children.append(ch)
            continue

        run = ch
        parts = []
        for t in run.findall(".//w:t", NS):
            if t.text:
                parts.append(t.text)
        run_text = "".join(parts)

        if not run_text:
            continue

        run_tokens = split_tokens(run_text)
        kept_tokens = []

        for tok in run_tokens:
            if updated_keep_counter[tok] > 0:
                kept_tokens.append(tok)
                updated_keep_counter[tok] -= 1
            # else drop tok

        kept_text = "".join(kept_tokens)

        if kept_text == "":
            continue

        # rebuild run with kept_text
        new_run = deepcopy(run)
        for node in list(new_run):
            if node.tag == w_tag("t"):
                new_run.remove(node)

        t_new = ET.Element(w_tag("t"))
        t_new.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_new.text = kept_text
        new_run.append(t_new)

        new_children.append(new_run)

    # Replace paragraph children
    for ch in list(p):
        p.remove(ch)
    for ch in new_children:
        p.append(ch)

# ----------------------------
# Step 2: apply explicit INS/DEL markers to filtered doc paragraph
# ----------------------------
def apply_delete_span(old_children, run_map, new_children, child_idx_ref, start_off, del_text, author, next_id, date):
    """
    Replace the original normal run span with a single <w:del>,
    ensuring deletion text does NOT remain as normal runs.
    """
    end_off = start_off + len(del_text)

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

    # fallback if cannot find
    if child_idx_ref[0] >= len(old_children) or old_children[child_idx_ref[0]] is not first_run:
        new_children.append(make_del(first_run, del_text, next_id, author, date))
        return next_id + 1

    # local offsets
    first_local = start_off - first_rm["start"]
    last_local = end_off - last_rm["start"]

    # split first run at start
    left_first, right_first = split_run_at(first_run, first_local)

    # split last run at end
    if first_run is last_run:
        _mid, right_last = split_run_at(right_first, last_local - first_local)
    else:
        _mid, right_last = split_run_at(last_run, last_local)

    # consume first_run
    child_idx_ref[0] += 1

    # skip until after last_run
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

def apply_markers_to_filtered_doc(p: ET.Element, updated_marked: str, author: str, next_id: int):
    """
    After filtering paragraph, apply:
      - <DEL:...> : wrap and remove span
      - <INS:...> : insert at correct place (next anchor strategy)
    """
    original_text, run_map = build_paragraph_fulltext_and_runs(p)
    if not original_text or not run_map:
        return next_id

    upd_tokens = tokenize_updated_atomic(updated_marked)
    cursor_char = 0
    events = []
    date = now_w3c()

    i = 0
    while i < len(upd_tokens):
        kind, payload = upd_tokens[i]

        if kind == "TEXTTOK":
            if original_text.startswith(payload, cursor_char):
                cursor_char += len(payload)
            # else ignore mismatch outside marker
            i += 1
            continue

        if kind == "DELBLOCK":
            if payload:
                pos = original_text.find(payload, cursor_char)
                if pos >= 0:
                    events.append((pos, "DEL", payload))
                    cursor_char = pos + len(payload)
            i += 1
            continue

        if kind == "INSBLOCK":
            ins_text = payload

            # next anchor: find next meaningful TEXT token
            j = i + 1
            anchor = None
            while j < len(upd_tokens):
                k2, p2 = upd_tokens[j]
                if k2 == "TEXTTOK" and p2.strip() != "":
                    anchor = p2
                    break
                j += 1

            if anchor:
                pos = original_text.find(anchor, cursor_char)
                insert_pos = pos if pos >= 0 else cursor_char
            else:
                insert_pos = len(original_text)

            if ins_text:
                events.append((insert_pos, "INS", ins_text))

            i += 1
            continue

        i += 1

    if not events:
        return next_id

    events.sort(key=lambda x: (x[0], 0 if x[1] == "DEL" else 1))

    old_children = list(p)
    new_children = []
    child_idx_ref = [0]

    for off, kind, text in events:
        if kind == "DEL":
            next_id = apply_delete_span(
                old_children, run_map, new_children, child_idx_ref,
                off, text, author, next_id, date
            )
        else:
            rm = find_run_covering_offset(run_map, off)
            if rm is None:
                continue
            base_run = rm["r"]

            # copy children until base_run
            while child_idx_ref[0] < len(old_children):
                ch = old_children[child_idx_ref[0]]
                if ch is base_run:
                    break
                new_children.append(ch)
                child_idx_ref[0] += 1

            # fallback: append
            if child_idx_ref[0] >= len(old_children) or old_children[child_idx_ref[0]] is not base_run:
                new_children.append(make_ins(base_run, text, next_id, author, date))
                next_id += 1
                continue

            local = off - rm["start"]
            local = max(0, min(local, rm["end"] - rm["start"]))

            left_run, right_run = split_run_at(base_run, local)

            # consume base_run
            child_idx_ref[0] += 1

            if left_run.findall(".//w:t", NS):
                new_children.append(left_run)

            new_children.append(make_ins(base_run, text, next_id, author, date))
            next_id += 1

            if right_run.findall(".//w:t", NS):
                new_children.append(right_run)

    # append remaining original children
    while child_idx_ref[0] < len(old_children):
        new_children.append(old_children[child_idx_ref[0]])
        child_idx_ref[0] += 1

    # replace paragraph children
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

        # OUTSIDE TEXT tokens
        updated_keep_text = strip_markers_keep_text(updated_marked)

        # ALSO include deletion tokens so filter doesn't drop them before wrapping
        del_texts = extract_all_del_text(updated_marked)
        del_keep_text = "".join(del_texts)

        keep_counter = Counter(split_tokens(updated_keep_text + del_keep_text))

        # 1) Filter doc paragraph tokens (drop words not present in updated or DEL blocks)
        filter_paragraph_by_updated_text(p, keep_counter)

        # 2) Apply markers INS/DEL
        next_id = apply_markers_to_filtered_doc(p, updated_marked, author, next_id)

        changed += 1

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Updated {changed} paragraphs (doc truth + drop missing + track changes).")

if __name__ == "__main__":
    apply_track_changes("word/document.xml", "updated.txt", author="Tanmay")
