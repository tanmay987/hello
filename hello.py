import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from collections import Counter
from datetime import datetime, timezone
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict

# ============================
# WordprocessingML namespace
# ============================
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ET.register_namespace("w", NS["w"])


def w_tag(tag: str) -> str:
    """Return fully qualified WordprocessingML tag for ElementTree."""
    return f"{{{NS['w']}}}{tag}"


def now_w3c() -> str:
    """Return current UTC timestamp in Word track-changes format."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ============================
# updated.txt parsing
# ============================
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")
MARK_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)

# Tokenizer:
# - words: letters/digits/underscore
# - punctuation: every other non-space char
# - spaces ignored as tokens
TOKEN_RE = re.compile(r"\w+|[^\w\s]", re.UNICODE)


def parse_updated_txt(path: str) -> Dict[int, str]:
    """
    Parse updated.txt into:
      { paragraph_index (1-based) : paragraph_string_with_markers }
    """
    out: Dict[int, str] = {}
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


def tokenize_words_punct(s: str) -> List[str]:
    """
    Tokenize into word + punctuation tokens.
    Spaces are not tokens.
    Examples:
      "school," -> ["school", ","]
      "in-car"  -> ["in", "-", "car"]
    """
    return [m.group(0) for m in TOKEN_RE.finditer(s)]


@dataclass
class UpdatedPlan:
    """
    Parsed representation of one updated paragraph.

    keep_tokens:
      tokens outside INS/DEL (what remains visible normal text)
    del_blocks:
      list of deletion blocks, each a token sequence
    ins_blocks:
      list of insertion blocks (raw string as-is)
    ins_anchors:
      for each ins block, the next KEEP token after it (or None if insertion at end)
      used to place insertion in the rebuilt paragraph.
    """
    keep_tokens: List[str]
    del_blocks: List[List[str]]
    ins_blocks: List[str]
    ins_anchors: List[Optional[str]]


def parse_updated_paragraph(updated_marked: str) -> UpdatedPlan:
    """
    Parse one updated paragraph into KEEP tokens + DEL token blocks + INS blocks.

    INS blocks remain raw text (we will insert as ONE <w:ins>).
    DEL blocks become token sequences for matching against document.xml tokens.

    Also computes insertion anchors:
      for each INS block, anchor = next KEEP token after that INS marker.
    """
    keep_tokens: List[str] = []
    del_blocks: List[List[str]] = []
    ins_blocks: List[str] = []
    ins_anchors: List[Optional[str]] = []

    # We also build a stream of items in order:
    # ("KEEP", token) or ("INS", raw) or ("DEL", token_list)
    stream: List[Tuple[str, object]] = []

    pos = 0
    for m in MARK_RE.finditer(updated_marked):
        outside = updated_marked[pos:m.start()]
        for t in tokenize_words_punct(outside):
            keep_tokens.append(t)
            stream.append(("KEEP", t))

        kind = m.group(1)
        inner = m.group(2)

        if kind == "INS":
            ins_blocks.append(inner)
            stream.append(("INS", inner))
        else:
            del_toks = tokenize_words_punct(inner)
            del_blocks.append(del_toks)
            stream.append(("DEL", del_toks))

        pos = m.end()

    tail = updated_marked[pos:]
    for t in tokenize_words_punct(tail):
        keep_tokens.append(t)
        stream.append(("KEEP", t))

    # Compute INS anchors: next KEEP token in the stream after INS
    for idx, (k, v) in enumerate(stream):
        if k != "INS":
            continue
        anchor = None
        for j in range(idx + 1, len(stream)):
            if stream[j][0] == "KEEP":
                anchor = stream[j][1]  # token string
                break
        ins_anchors.append(anchor)

    return UpdatedPlan(
        keep_tokens=keep_tokens,
        del_blocks=del_blocks,
        ins_blocks=ins_blocks,
        ins_anchors=ins_anchors,
    )


# ============================
# document.xml tokenization with run mapping
# ============================
@dataclass
class DocToken:
    """
    One token from document.xml paragraph.
    token: token string (word/punctuation)
    run_template: the original <w:r> node this token came from (formatting source)
    """
    token: str
    run_template: ET.Element


def extract_run_text(run: ET.Element) -> str:
    """Concatenate visible text from all <w:t> nodes inside this run."""
    parts = []
    for t in run.findall(".//w:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def build_doc_tokens(p: ET.Element) -> List[DocToken]:
    """
    Convert document.xml paragraph into token list with formatting mapping.

    We walk runs in order:
      each token inherits formatting from the run it came from.
    """
    tokens: List[DocToken] = []
    for r in p.findall(".//w:r", NS):
        txt = extract_run_text(r)
        if not txt:
            continue
        for tok in tokenize_words_punct(txt):
            tokens.append(DocToken(tok, r))
    return tokens


# ============================
# Track-change node constructors
# ============================
def clone_run_with_text(run_template: ET.Element, text: str) -> ET.Element:
    """
    Clone run template and set its visible text to `text`
    with xml:space="preserve" to avoid space-loss collisions.
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


def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str) -> ET.Element:
    """Create one <w:ins> containing the whole inserted text."""
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)
    ins.append(clone_run_with_text(run_template, text))
    return ins


def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str) -> ET.Element:
    """Create one <w:del> containing the whole deleted text."""
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
# Core matching helpers
# ============================
def find_sequence(tokens: List[str], seq: List[str], start: int = 0) -> Optional[int]:
    """
    Find seq inside tokens starting from index start.
    Return starting index or None.
    """
    if not seq:
        return None
    n = len(tokens)
    m = len(seq)
    for i in range(start, n - m + 1):
        if tokens[i : i + m] == seq:
            return i
    return None


# ============================
# Paragraph rebuild
# ============================
def rebuild_paragraph(
    p: ET.Element,
    updated: UpdatedPlan,
    author: str,
    next_id: int,
) -> int:
    """
    Build a NEW paragraph content using:
      - document.xml as formatting ground truth
      - updated.txt as token truth for which tokens remain
      - apply <DEL> as <w:del>
      - apply <INS> as <w:ins>

    Also removes doc tokens not present in updated KEEP tokens.

    Returns updated next_id.
    """
    date = now_w3c()

    # ---------- 1) Read doc tokens ----------
    doc_tokens = build_doc_tokens(p)
    doc_token_strs = [dt.token for dt in doc_tokens]

    # ---------- 2) Determine what stays visible ----------
    keep_counter = Counter(updated.keep_tokens)

    # We'll build filtered normal tokens from doc, preserving order.
    # Each normal token still carries its run template.
    normal_tokens: List[DocToken] = []
    for dt in doc_tokens:
        if keep_counter[dt.token] > 0:
            normal_tokens.append(dt)
            keep_counter[dt.token] -= 1
        else:
            # token not in updated KEEP => removed
            pass

    # ---------- 3) Build a paragraph child list ----------
    # Keep paragraph properties (<w:pPr>) as-is.
    pPr = p.find("./w:pPr", NS)

    new_children: List[ET.Element] = []
    if pPr is not None:
        new_children.append(deepcopy(pPr))

    # We'll create a working token list (normal stream),
    # but deletions/insertions will be injected while building XML.
    normal_strs = [t.token for t in normal_tokens]

    # ---------- 4) Apply deletions (wrap+remove from normal) ----------
    # We will delete token sequences from the *normal stream* when present,
    # and insert <w:del> nodes in their place.
    #
    # If a DEL block is not found in normal stream (because removed already),
    # we search in original doc token stream to still show it as deletion.
    #
    # We'll process DEL blocks in order.
    cursor = 0
    del_nodes: List[Tuple[int, ET.Element, int]] = []  # (insert_index, node, token_count_removed)

    for del_seq in updated.del_blocks:
        if not del_seq:
            continue

        found = find_sequence(normal_strs, del_seq, start=cursor)

        if found is not None:
            # Build del text as it appears (join tokens without extra spaces).
            # We keep it simple: join with single space between WORD tokens if needed is complicated.
            # Instead, emit exactly " ".join(words) for readability.
            del_text = " ".join(del_seq)

            run_template = normal_tokens[found].run_template
            node = make_del(run_template, del_text, next_id, author, date)
            next_id += 1

            del_nodes.append((found, node, len(del_seq)))

            # Remove that seq from normal tokens/strs
            del normal_tokens[found : found + len(del_seq)]
            del normal_strs[found : found + len(del_seq)]

            cursor = found  # continue from deletion point
        else:
            # Not found in normal: try from original doc tokens to still show
            found2 = find_sequence(doc_token_strs, del_seq, start=0)
            if found2 is None:
                continue

            del_text = " ".join(del_seq)
            run_template = doc_tokens[found2].run_template
            node = make_del(run_template, del_text, next_id, author, date)
            next_id += 1

            # place near end if we cannot map
            del_nodes.append((len(normal_tokens), node, 0))

    # After processing, normal_tokens already updated.

    # ---------- 5) Recompute normal_strs after deletions ----------
    normal_strs = [t.token for t in normal_tokens]

    # ---------- 6) Apply insertions based on anchor token ----------
    # For each INS block:
    #   - find the anchor token occurrence in normal_strs
    #   - insert <w:ins> before anchor
    #   - if no anchor => append at end
    ins_nodes: List[Tuple[int, ET.Element]] = []
    for ins_text, anchor in zip(updated.ins_blocks, updated.ins_anchors):
        if ins_text is None:
            continue
        ins_text = ins_text.strip("\n")
        if ins_text == "":
            continue

        if anchor is None:
            idx = len(normal_tokens)
            run_template = normal_tokens[-1].run_template if normal_tokens else (doc_tokens[0].run_template if doc_tokens else ET.Element(w_tag("r")))
            ins_nodes.append((idx, make_ins(run_template, ins_text, next_id, author, date)))
            next_id += 1
            continue

        # find anchor in remaining normal tokens
        try:
            idx = normal_strs.index(anchor)
        except ValueError:
            idx = len(normal_tokens)

        run_template = normal_tokens[idx].run_template if idx < len(normal_tokens) else (normal_tokens[-1].run_template if normal_tokens else doc_tokens[0].run_template)
        ins_nodes.append((idx, make_ins(run_template, ins_text, next_id, author, date)))
        next_id += 1

    # ---------- 7) Merge all events into one stream ----------
    # We'll insert del_nodes and ins_nodes into the normal token stream positions.
    #
    # We'll build an "events map": position -> list of nodes
    event_map: Dict[int, List[ET.Element]] = {}

    for pos, node, _removed in del_nodes:
        event_map.setdefault(pos, []).append(node)

    for pos, node in ins_nodes:
        event_map.setdefault(pos, []).append(node)

    # ---------- 8) Emit runs while preserving formatting ----------
    # We emit normal tokens grouped by run_template (formatting),
    # and inject INS/DEL nodes at positions.
    def flush_run_group(run_template: ET.Element, buf: List[str]):
        if not buf:
            return
        # simple join with space between tokens is not correct in general.
        # But our tokenizer removed spaces, so we need a "smart join".
        #
        # Smart join: if next token is punctuation, don't add space before it.
        out = []
        for t in buf:
            if not out:
                out.append(t)
            else:
                if re.match(r"[^\w\s]", t):  # punctuation token
                    out.append(t)
                else:
                    out.append(" " + t)
        text = "".join(out)

        new_children.append(clone_run_with_text(run_template, text))
        buf.clear()

    buf: List[str] = []
    current_template: Optional[ET.Element] = None

    # Walk positions 0..len(normal_tokens) and inject events
    for i in range(len(normal_tokens) + 1):
        # inject events at i
        if i in event_map:
            flush_run_group(current_template, buf)
            # stable ordering: DEL nodes first then INS nodes (Word typical)
            for node in event_map[i]:
                new_children.append(node)

        if i == len(normal_tokens):
            break

        dt = normal_tokens[i]
        if current_template is None:
            current_template = dt.run_template

        if dt.run_template is not current_template:
            flush_run_group(current_template, buf)
            current_template = dt.run_template

        buf.append(dt.token)

    flush_run_group(current_template, buf)

    # ---------- 9) Replace paragraph content ----------
    for ch in list(p):
        p.remove(ch)
    for ch in new_children:
        p.append(ch)

    return next_id


# ============================
# Main: rebuild whole document.xml
# ============================
def rebuild_document_xml(
    document_xml_path: str,
    updated_txt_path: str,
    out_xml_path: str,
    author: str = "Tanmay",
):
    """
    Create a NEW document xml with tracked changes based on updated.txt.

    Paragraph mapping:
      updated.txt uses [P1] .. [Pn] mapped to Word paragraphs (1-based).
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
        updated_marked = edits[i]
        plan = parse_updated_paragraph(updated_marked)
        next_id = rebuild_paragraph(p, plan, author, next_id)
        changed += 1

    tree.write(out_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Created new XML: {out_xml_path}")
    print(f"✅ Updated {changed} paragraphs using track changes.")


if __name__ == "__main__":
    rebuild_document_xml(
        document_xml_path="word/document.xml",
        updated_txt_path="updated.txt",
        out_xml_path="word/document_updated.xml",
        author="Tanmay",
    )
