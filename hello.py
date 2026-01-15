import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime, timezone

# ============================
# WordprocessingML Namespace
# ============================
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ET.register_namespace("w", NS["w"])


def w_tag(tag: str) -> str:
    """Return a fully-qualified WordprocessingML tag for ElementTree."""
    return f"{{{NS['w']}}}{tag}"


def now_w3c() -> str:
    """Return current UTC time formatted for Word track changes."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ============================
# updated.txt parsing
# ============================
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")
MARK_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)
WORD_RE = re.compile(r"\s+|[^\s]+")  # tokenizes into ["word", " ", "word", ...]


def parse_updated_txt(path: str) -> dict[int, str]:
    """Parse updated.txt into {paragraph_id: marked_content}."""
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


def split_tokens_preserve_spaces(text: str) -> list[str]:
    """Split into word/space tokens while preserving spaces as tokens."""
    return [m.group(0) for m in WORD_RE.finditer(text)]


def tokenize_updated_with_markers(s: str) -> list[tuple[str, str]]:
    """
    Convert updated marked text into token stream:
      ("TEXT", token) = token outside markers
      ("INS", token)  = token inside <INS:...>
      ("DEL", token)  = token inside <DEL:...>
    """
    tokens: list[tuple[str, str]] = []
    pos = 0

    for m in MARK_RE.finditer(s):
        outside = s[pos:m.start()]
        for t in split_tokens_preserve_spaces(outside):
            tokens.append(("TEXT", t))

        kind = m.group(1)          # "INS" or "DEL"
        inner = m.group(2)
        for t in split_tokens_preserve_spaces(inner):
            tokens.append((kind, t))

        pos = m.end()

    tail = s[pos:]
    for t in split_tokens_preserve_spaces(tail):
        tokens.append(("TEXT", t))

    return tokens


# ============================
# document.xml paragraph text + run map
# ============================
def build_paragraph_fulltext_and_runs(p: ET.Element):
    """
    Return (full_text, runs) for ALL runs inside paragraph.

    runs = list of dict:
      {
        "r": run_element,
        "text": run_text,
        "start": global_start_char,
        "end": global_end_char
      }

    We use .//w:r so it includes runs inside hyperlinks etc.
    """
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
    """Find the run-map entry that covers char_off."""
    if not run_map:
        return None
    if char_off < 0:
        return run_map[0]
    if char_off >= run_map[-1]["end"]:
        return run_map[-1]
    for m in run_map:
        if m["start"] <= char_off < m["end"]:
            return m
    return run_map[-1]


def split_run_at(run: ET.Element, offset: int):
    """
    Split a run into two runs at offset (in run's combined text).
    Preserves formatting by deepcopy cloning the run and rebuilding w:t.
    """
    # Collect run's visible text from w:t
    t_nodes = run.findall(".//w:t", NS)
    combined = "".join([(t.text or "") for t in t_nodes])

    left_text = combined[:offset]
    right_text = combined[offset:]

    left = deepcopy(run)
    right = deepcopy(run)

    # Remove old w:t nodes
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
        # Word requires xml:space="preserve" to keep leading/trailing spaces
        if txt.startswith(" ") or txt.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = txt
        run_elem.append(t)

    add_wt(left, left_text)
    add_wt(right, right_text)
    return left, right


# ============================
# Track change nodes
# ============================
def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str):
    """Create <w:del> wrapper with <w:delText>."""
    dele = ET.Element(w_tag("del"))
    dele.set(w_tag("id"), str(del_id))
    dele.set(w_tag("author"), author)
    dele.set(w_tag("date"), date)

    r = deepcopy(run_template)

    # remove normal <w:t>, replace with <w:delText>
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    dt = ET.Element(w_tag("delText"))
    if text.startswith(" ") or text.endswith(" "):
        dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    dt.text = text
    r.append(dt)

    dele.append(r)
    return dele


def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str):
    """Create <w:ins> wrapper with <w:t>."""
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)

    r = deepcopy(run_template)

    # remove existing <w:t> nodes
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    t = ET.Element(w_tag("t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)

    ins.append(r)
    return ins


# ============================
# Core: apply markers (exact placement)
# ============================
def apply_only_markers_wordwise_full_rebuild(
    p: ET.Element,
    updated_marked: str,
    author: str,
    next_id: int,
    del_lookahead: int = 30,
):
    """
    Apply policy:
      - Keep document.xml normal text intact
      - Ignore any outside-marker differences
      - Only apply <INS:...> and <DEL:...> markers
      - Insert <w:ins>/<w:del> at correct character offsets inside paragraph

    How we do it:
      1) Parse original paragraph full text (document.xml authoritative)
      2) Tokenize original into word/space tokens (doc_tokens)
      3) Tokenize updated_marked into stream of TEXT/INS/DEL tokens
      4) Walk updated tokens and maintain doc pointer (doc_i) and char cursor
         - TEXT tokens: advance only if token equals current doc token, else ignore
         - INS tokens: insert <w:ins> at current char cursor (does not consume doc)
         - DEL tokens: try to match next doc token == del token, else lookahead resync; if found delete there
      5) We collect a list of "events" to apply: insertions/deletions at char offsets
      6) Apply events by rebuilding paragraph runs using splits + injecting nodes
    """
    original_text, run_map = build_paragraph_fulltext_and_runs(p)
    if not original_text or not run_map:
        return next_id  # nothing to do

    # Word+spaces tokens for authoritative doc
    doc_tokens = split_tokens_preserve_spaces(original_text)

    # Build token -> char offset map (start offset of each token)
    doc_tok_offsets = []
    cur = 0
    for tok in doc_tokens:
        doc_tok_offsets.append(cur)
        cur += len(tok)

    # Tokenize updated marked stream
    upd_tokens = tokenize_updated_with_markers(updated_marked)

    # Walk pointers
    doc_i = 0
    cursor_char = 0

    # Events: (char_offset, "INS"/"DEL", text)
    # DEL text consumes doc tokens; INS does not
    events: list[tuple[int, str, str]] = []

    for kind, tok in upd_tokens:
        if kind == "INS":
            if tok:
                events.append((cursor_char, "INS", tok))
            continue

        if kind == "DEL":
            if not tok:
                continue

            # Ideal: delete exactly when doc token matches
            if doc_i < len(doc_tokens) and doc_tokens[doc_i] == tok:
                events.append((cursor_char, "DEL", tok))
                cursor_char += len(tok)
                doc_i += 1
                continue

            # Lookahead resync: find tok in next N tokens
            found = -1
            for j in range(doc_i, min(len(doc_tokens), doc_i + del_lookahead)):
                if doc_tokens[j] == tok:
                    found = j
                    break

            if found >= 0:
                # advance doc to found (unchanged doc is kept)
                while doc_i < found:
                    cursor_char += len(doc_tokens[doc_i])
                    doc_i += 1

                # now delete
                events.append((cursor_char, "DEL", tok))
                cursor_char += len(tok)
                doc_i += 1
            else:
                # cannot locate => ignore deletion (fallback to doc)
                continue
            continue

        # kind == "TEXT"
        if not tok:
            continue

        # Only advance cursor if TEXT token matches document token
        if doc_i < len(doc_tokens) and doc_tokens[doc_i] == tok:
            cursor_char += len(tok)
            doc_i += 1
        else:
            # mismatch => ignore (document authoritative)
            continue

    if not events:
        return next_id

    # ----------------------------
    # APPLY EVENTS to paragraph by rebuilding runs using splits
    # ----------------------------
    # Sort events by offset; for same offset, apply DEL first then INS (Word semantics)
    events.sort(key=lambda x: (x[0], 0 if x[1] == "DEL" else 1))

    # Paragraph direct children (may include non-run nodes)
    old_children = list(p)

    # We'll rebuild children into new_children
    new_children = []
    child_idx = 0

    # Helper: copy children until we reach a particular run element
    def copy_children_until(target_run: ET.Element):
        nonlocal child_idx
        while child_idx < len(old_children):
            ch = old_children[child_idx]
            if ch is target_run:
                break
            new_children.append(ch)
            child_idx += 1

    # We will keep an evolving "current run_map" as we split.
    # Approach: apply events one-by-one by:
    #   - find run at event offset
    #   - copy everything before it
    #   - split that run into left/right at local offset
    #   - add left, inject change, add right
    #   - continue
    #
    # For correctness with multiple events, we must track offset shifts due to insertions/deletions.
    # But IMPORTANT: our policy says doc text remains intact, even deletions keep text (as <w:del>),
    # so the visible underlying text *still exists in XML* and offsets are stable if we rebuild carefully.
    #
    # We apply in original coordinate system by progressively rebuilding.

    date = now_w3c()

    # We'll rebuild using "active text stream" that is still original_text.
    # But once we split, run_map changes. We'll handle by searching by char offset
    # against original run_map start/end (works if we always split based on original).
    #
    # Implementation strategy:
    #   Use original run_map to locate target run each time.
    #   Compute local offset in that run.
    #
    # This is OK because run_map references original nodes; once we append split clones,
    # the original node should be consumed and not reused again. So we also need a guard
    # to avoid targeting same original run after it was replaced.
    consumed_runs = set()

    for (off, kind, text) in events:
        # Find template run at this offset
        rm = find_run_at_offset(run_map, off)
        if rm is None:
            continue
        base_run = rm["r"]

        # If base_run already consumed (split before), find nearest previous/next not consumed
        if base_run in consumed_runs:
            # fallback: walk run_map to find a run not consumed that still covers or follows offset
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
            if candidate is None:
                candidate = run_map[-1]
            rm = candidate
            base_run = rm["r"]

        # Copy all children until reaching this run element
        copy_children_until(base_run)

        # Now child at child_idx should be base_run (or we've reached end)
        if child_idx >= len(old_children) or old_children[child_idx] is not base_run:
            # Cannot place precisely => append change node at end of paragraph
            template = rm["r"]
            if kind == "INS":
                new_children.append(make_ins(template, text, next_id, author, date))
                next_id += 1
            else:
                new_children.append(make_del(template, text, next_id, author, date))
                next_id += 1
            continue

        # Compute local offset inside run
        local = off - rm["start"]
        local = max(0, min(local, rm["end"] - rm["start"]))

        # Split base run into left/right at local
        left_run, right_run = split_run_at(base_run, local)

        # Consume the base_run child
        child_idx += 1
        consumed_runs.add(base_run)

        # Append left part
        if left_run.findall(".//w:t", NS):
            new_children.append(left_run)

        # Inject track change node
        if kind == "INS":
            new_children.append(make_ins(base_run, text, next_id, author, date))
            next_id += 1
        else:
            new_children.append(make_del(base_run, text, next_id, author, date))
            next_id += 1

        # Append right part
        if right_run.findall(".//w:t", NS):
            new_children.append(right_run)

        # Important: right_run is a clone, but old_children still continues from after base_run.
        # We do NOT attempt to merge it back; this is okay because we are rebuilding paragraph children.

    # Append any remaining untouched children
    while child_idx < len(old_children):
        new_children.append(old_children[child_idx])
        child_idx += 1

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
    """
    Apply explicit marker-based track changes document-wide.

    Input:
      - document_xml_path: path to extracted 'word/document.xml'
      - updated_txt_path: updated.txt with lines like:
            [P1] Hello <DEL:world><INS:Tanmay>
            [P2] ...
    Policy:
      - Ignore all differences outside markers
      - Apply only <INS>/<DEL> as track changes
    """
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

        next_id = apply_only_markers_wordwise_full_rebuild(
            p=p,
            updated_marked=marked,
            author=author,
            next_id=next_id,
        )
        changed += 1

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Applied marker-only tracked changes to {changed} paragraphs.")


if __name__ == "__main__":
    # Adjust if needed
    apply_track_changes("word/document.xml", "updated.txt", author="Tanmay")
