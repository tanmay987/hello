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
    """Return fully-qualified WordprocessingML tag."""
    return f"{{{NS['w']}}}{tag}"

def now_w3c() -> str:
    """UTC timestamp in Word track-changes format."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

# ============================
# updated.txt parsing
# ============================
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")
MARK_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)

# Used ONLY for aligning "TEXT outside markers" to doc text.
# We preserve spaces tokens too.
WORD_RE = re.compile(r"\s+|[^\s]+")

def parse_updated_txt(path: str) -> dict[int, str]:
    """Load updated.txt lines like [P12] ... into a dict."""
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
    """Tokenize into word/space tokens (spaces preserved)."""
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
# Word paragraph: text + run map
# ============================
def build_paragraph_fulltext_and_runs(p: ET.Element):
    """
    Extract visible paragraph text and run map for ALL runs (.//w:r).

    Returns:
      full_text: concatenated string of all run texts
      run_map: list of dict entries:
        {"r": run_elem, "start": int, "end": int, "text": run_text}
    """
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
    """Return run_map entry covering character offset off."""
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
    Split a run at offset into (left,right) preserving formatting.

    CRITICAL:
      ALWAYS apply xml:space="preserve" so boundary spaces never get eaten.
    """
    t_nodes = run.findall(".//w:t", NS)
    combined = "".join([(t.text or "") for t in t_nodes])

    left_text = combined[:offset]
    right_text = combined[offset:]

    left = deepcopy(run)
    right = deepcopy(run)

    # Remove direct w:t children (common case)
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
# Track change nodes (atomic)
# ============================
def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str):
    """Create <w:ins> with ONE run containing the entire inserted text."""
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)

    r = deepcopy(run_template)

    # remove normal <w:t> children
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
    """Create <w:del> with ONE run containing the entire deleted text."""
    dele = ET.Element(w_tag("del"))
    dele.set(w_tag("id"), str(del_id))
    dele.set(w_tag("author"), author)
    dele.set(w_tag("date"), date)

    r = deepcopy(run_template)

    # remove normal <w:t> children
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
# Apply delete span CORRECTLY (remove from normal runs)
# ============================
def apply_delete_span(
    old_children: list[ET.Element],
    run_map: list[dict],
    new_children: list[ET.Element],
    child_idx_ref: list[int],
    consumed_runs: set,
    start_off: int,
    del_text: str,
    author: str,
    next_id: int,
    date: str,
):
    """
    Replace the original visible text span [start_off, start_off+len(del_text))
    with ONE <w:del> node and remove original normal text from paragraph.
    """

    end_off = start_off + len(del_text)

    first_rm = find_run_covering_offset(run_map, start_off)
    last_rm = find_run_covering_offset(run_map, max(start_off, end_off - 1))

    if not first_rm or not last_rm:
        return next_id

    first_run = first_rm["r"]
    last_run = last_rm["r"]

    # Copy children until first_run
    child_idx = child_idx_ref[0]
    while child_idx < len(old_children):
        ch = old_children[child_idx]
        if ch is first_run:
            break
        new_children.append(ch)
        child_idx += 1
    child_idx_ref[0] = child_idx

    # If cannot locate first_run precisely, fallback: just insert del node
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

    # Consume first_run
    child_idx_ref[0] += 1
    consumed_runs.add(first_run)

    # Skip original children until after last_run
    if first_run is not last_run:
        while child_idx_ref[0] < len(old_children):
            ch = old_children[child_idx_ref[0]]
            child_idx_ref[0] += 1
            if ch is last_run:
                consumed_runs.add(last_run)
                break

    # Append kept left part
    if left_first.findall(".//w:t", NS):
        new_children.append(left_first)

    # Append deletion node (ONLY place deleted text exists)
    new_children.append(make_del(first_run, del_text, next_id, author, date))
    next_id += 1

    # Append kept right part
    if right_last.findall(".//w:t", NS):
        new_children.append(right_last)

    return next_id

# ============================
# Apply markers to paragraph
# ============================
def apply_markers_to_paragraph(p: ET.Element, updated_marked: str, author: str, next_id: int):
    """
    Correct behavior:
      - INS block inserted as one node
      - DEL block wraps and REMOVES original text span (no duplicates)
      - Outside marker differences ignored
    """
    original_text, run_map = build_paragraph_fulltext_and_runs(p)
    if not original_text or not run_map:
        return next_id

    # Tokenize doc (for TEXT alignment only)
    doc_tokens = split_text_tokens(original_text)
    doc_i = 0
    cursor_char = 0

    upd_tokens = tokenize_updated_atomic(updated_marked)

    # Events: (offset, kind, text)
    # DEL event means delete span from original.
    events: list[tuple[int, str, str]] = []

    for kind, payload in upd_tokens:
        if kind == "INSBLOCK":
            if payload:
                events.append((cursor_char, "INS", payload))
            continue

        if kind == "DELBLOCK":
            del_text = payload
            if not del_text:
                continue

            # Find where this deletion exists in original paragraph text
            # starting from cursor_char. This ensures we delete correct span.
            pos = original_text.find(del_text, cursor_char)
            if pos >= 0:
                events.append((pos, "DEL", del_text))
                cursor_char = pos + len(del_text)
                continue

            # If exact substring not found, ignore deletion (fallback to doc)
            continue

        # TEXTTOK
        tok = payload
        if doc_i < len(doc_tokens) and doc_tokens[doc_i] == tok:
            cursor_char += len(tok)
            doc_i += 1
        else:
            # ignore mismatches outside markers
            continue

    if not events:
        return next_id

    # Apply DEL before INS if same offset
    events.sort(key=lambda x: (x[0], 0 if x[1] == "DEL" else 1))

    old_children = list(p)
    new_children: list[ET.Element] = []
    child_idx_ref = [0]
    consumed_runs = set()
    date = now_w3c()

    for off, kind, text in events:
        if kind == "DEL":
            next_id = apply_delete_span(
                old_children=old_children,
                run_map=run_map,
                new_children=new_children,
                child_idx_ref=child_idx_ref,
                consumed_runs=consumed_runs,
                start_off=off,
                del_text=text,
                author=author,
                next_id=next_id,
                date=date,
            )
        else:
            # INS: insert at exact position
            rm = find_run_covering_offset(run_map, off)
            if rm is None:
                continue
            base_run = rm["r"]

            # Copy children until base_run
            while child_idx_ref[0] < len(old_children):
                ch = old_children[child_idx_ref[0]]
                if ch is base_run:
                    break
                new_children.append(ch)
                child_idx_ref[0] += 1

            # If can't place precisely, append
            if child_idx_ref[0] >= len(old_children) or old_children[child_idx_ref[0]] is not base_run:
                new_children.append(make_ins(base_run, text, next_id, author, date))
                next_id += 1
                continue

            # Split base run at insertion point, then insert INS between left and right
            local = off - rm["start"]
            local = max(0, min(local, rm["end"] - rm["start"]))

            left_run, right_run = split_run_at(base_run, local)

            # consume base run
            child_idx_ref[0] += 1
            consumed_runs.add(base_run)

            if left_run.findall(".//w:t", NS):
                new_children.append(left_run)

            new_children.append(make_ins(base_run, text, next_id, author, date))
            next_id += 1

            if right_run.findall(".//w:t", NS):
                new_children.append(right_run)

    # append remaining children
    while child_idx_ref[0] < len(old_children):
        new_children.append(old_children[child_idx_ref[0]])
        child_idx_ref[0] += 1

    # Replace paragraph children
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

        next_id = apply_markers_to_paragraph(p, marked, author, next_id)
        changed += 1

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Applied tracked changes to {changed} paragraphs.")


if __name__ == "__main__":
    # Run on extracted docx folder:
    #   word/document.xml must exist
    # updated.txt must exist
    apply_track_changes("word/document.xml", "updated.txt", author="Tanmay")
