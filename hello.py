"""
build_documentxml_trackchanges_preserve_runs_spaces.py

WHAT THIS SCRIPT DOES
---------------------
Creates a NEW Word document.xml (WordprocessingML) that:

1) Uses updated.txt as "ground truth" for paragraph content.
   updated.txt contains special markers:
     - <INS: ...>  -> inserted text
     - <DEL: ...>  -> deleted text

2) Preserves formatting and run structure from the original document.xml.
   - We do NOT create a new run per word.
   - We reuse existing runs (<w:r>) as "buckets" and fill them sequentially.

3) Preserves spaces EXACTLY:
   - Unlike earlier versions, we tokenize whitespace as real tokens.
   - We never guess spaces.
   - We join tokens with ''.join() so spacing remains exactly as in updated.txt.

4) Applies Word Track Changes:
   - <INS: ...> becomes <w:ins>
   - <DEL: ...> becomes <w:del>

OUTPUT:
-------
- A new XML file: word/document_updated.xml

REQUIREMENTS:
-------------
- Your updated.txt must have lines like:
    [P1] This is text <INS: inserted> <DEL: deleted>
- Paragraph indexing is 1-based and matches Word paragraph order.
"""

import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime, timezone
from dataclasses import dataclass
from typing import List, Dict, Union


# ============================
# WordprocessingML namespace
# ============================

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ET.register_namespace("w", NS["w"])


def w_tag(tag: str) -> str:
    """
    Convert a WordprocessingML tag name into fully qualified XML tag.

    Example:
      w_tag("p") -> "{http://schemas.../main}p"

    Why:
      ElementTree needs fully-qualified tags for namespaced XML.
    """
    return f"{{{NS['w']}}}{tag}"


def now_w3c() -> str:
    """
    Return current UTC datetime in Word's track-changes timestamp format.

    Example:
      2026-01-21T12:34:56Z
    """
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ============================
# updated.txt parsing
# ============================

# Example line:
# [P12] some text <INS: x> <DEL: y>
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")

# Matches <INS: ...> or <DEL: ...> with any content inside.
MARK_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)

# ✅ SPACE-PRESERVING tokenization
# This captures:
#   1) whitespace sequences: \s+
#   2) word sequences: \w+
#   3) punctuation: [^\w\s]
TOKEN_RE = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)


def parse_updated_txt(path: str) -> Dict[int, str]:
    """
    Load updated.txt file and map paragraphs.

    Input file format:
      [P1] text...
      [P2] text...

    Returns:
      dict[int, str]
      where key is paragraph number (1-based) and value is marked paragraph string.
    """
    out: Dict[int, str] = {}
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")  # keep file clean
            if not line.strip():
                continue  # skip empty lines

            m = LINE_RE.match(line)
            if not m:
                continue  # skip malformed lines

            para_id = int(m.group(1))
            para_text = m.group(2)

            out[para_id] = para_text

    return out


def tokenize_with_spaces(s: str) -> List[str]:
    """
    Tokenize text into a sequence where spaces are preserved.

    This tokenization NEVER loses spacing information.

    Example:
      "Hello, world" =>
        ["Hello", ",", " ", "world"]

      "I  went" =>
        ["I", "  ", "went"]

    Why:
      Word formatting problems earlier came from losing whitespace.
      Now we preserve exact whitespace from updated.txt.
    """
    return [m.group(0) for m in TOKEN_RE.finditer(s)]


# ============================
# Updated paragraph plan objects
# ============================

@dataclass
class KeepTok:
    """
    Represents text outside <INS:...> and <DEL:...>.

    token can be:
      - whitespace like "   "
      - a word like "hello"
      - punctuation like ","
    """
    token: str


@dataclass
class InsBlock:
    """
    Represents an insertion block from updated.txt.

    This entire text will become ONE <w:ins> node.
    """
    text: str


@dataclass
class DelBlock:
    """
    Represents a deletion block from updated.txt.

    This entire text will become ONE <w:del> node.
    """
    text: str


# UpdatedItem is one of KeepTok / InsBlock / DelBlock
UpdatedItem = Union[KeepTok, InsBlock, DelBlock]


@dataclass
class UpdatedPlan:
    """
    Parsed version of one paragraph line from updated.txt.

    items preserves the exact order:
      KeepTok(...) KeepTok(...) InsBlock(...) KeepTok(...) DelBlock(...)

    We process these sequentially to build the output paragraph.
    """
    items: List[UpdatedItem]


def parse_updated_paragraph(updated_marked: str) -> UpdatedPlan:
    """
    Parse one updated paragraph (string) into UpdatedPlan.

    It walks through markers in the string:
      - tokens outside markers become KeepTok(token)
      - <INS: ...> becomes InsBlock(inner)
      - <DEL: ...> becomes DelBlock(inner)

    Output:
      UpdatedPlan(items=[...])
    """
    items: List[UpdatedItem] = []
    pos = 0

    for m in MARK_RE.finditer(updated_marked):
        # Everything before marker is outside text
        outside = updated_marked[pos:m.start()]
        for t in tokenize_with_spaces(outside):
            items.append(KeepTok(t))

        # Marker content
        kind = m.group(1)      # INS or DEL
        inner = m.group(2)     # inner text

        if kind == "INS":
            items.append(InsBlock(inner))
        else:
            items.append(DelBlock(inner))

        pos = m.end()

    # Add remaining tail outside markers
    tail = updated_marked[pos:]
    for t in tokenize_with_spaces(tail):
        items.append(KeepTok(t))

    return UpdatedPlan(items=items)


# ============================
# document.xml helpers
# ============================

def extract_run_text(run: ET.Element) -> str:
    """
    Extract visible text from a run (<w:r>).

    A run typically contains:
      <w:r>
        <w:rPr>...</w:rPr>      (formatting)
        <w:t>some text</w:t>    (visible text)
      </w:r>

    Runs can contain multiple <w:t>. We concatenate them.
    """
    parts = []
    for t in run.findall(".//w:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def clone_run_with_text(run_template: ET.Element, text: str) -> ET.Element:
    """
    Clone a run element to preserve formatting, then replace its text.

    - Deep-copy run_template (keeps <w:rPr>)
    - Remove existing <w:t>
    - Insert a new <w:t xml:space="preserve">text</w:t>

    Why xml:space="preserve":
      Without it, Word collapses leading/trailing spaces.
    """
    r = deepcopy(run_template)

    # Remove direct <w:t> children
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    # Add a new <w:t> if non-empty
    if text != "":
        t = ET.Element(w_tag("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        r.append(t)

    return r


def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str) -> ET.Element:
    """
    Create Word insertion node <w:ins> containing ONE run.

    Structure:
      <w:ins w:id=".." w:author=".." w:date="..">
        <w:r> ... <w:t>INSERTED</w:t> </w:r>
      </w:ins>
    """
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)

    ins.append(clone_run_with_text(run_template, text))
    return ins


def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str) -> ET.Element:
    """
    Create Word deletion node <w:del> containing ONE run.

    Structure:
      <w:del ...>
        <w:r>
          <w:delText>DELETED</w:delText>
        </w:r>
      </w:del>

    Note:
      deletion text is stored in <w:delText>, NOT <w:t>.
    """
    dele = ET.Element(w_tag("del"))
    dele.set(w_tag("id"), str(del_id))
    dele.set(w_tag("author"), author)
    dele.set(w_tag("date"), date)

    r = deepcopy(run_template)

    # Remove normal <w:t>
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
# Preserve runs using "token budget"
# ============================

@dataclass
class RunBucket:
    """
    One run from document.xml used as a formatting bucket.

    token_budget:
      how many tokens originally belonged to this run.
      We fill updated tokens into runs following this budget.

    Why:
      This preserves the *same run distribution* (formatting layout) as original.
    """
    run_template: ET.Element
    token_budget: int


def build_run_buckets(p: ET.Element) -> List[RunBucket]:
    """
    Build run buckets from paragraph direct runs (./w:r).

    Each run becomes a bucket with:
      - formatting template
      - token budget calculated from original run text tokenization

    Spaces are tokens, so spacing distribution also influences budgets.
    """
    buckets: List[RunBucket] = []

    for r in p.findall("./w:r", NS):
        txt = extract_run_text(r)
        toks = tokenize_with_spaces(txt)   # includes whitespace
        buckets.append(RunBucket(run_template=r, token_budget=len(toks)))

    return buckets


def rebuild_paragraph_preserve_runs(p: ET.Element, plan: UpdatedPlan, author: str, next_id: int) -> int:
    """
    Rebuild ONE paragraph using:
      - updated.txt tokens as content truth
      - document.xml runs as formatting buckets

    Steps:
      1) keep pPr if present
      2) reuse runs by filling them with updated tokens according to token budgets
      3) when encountering <INS>/<DEL>, close current run and insert track-change node
      4) preserve remaining runs as empty

    Returns:
      updated next_id for track-change ids.
    """
    date = now_w3c()

    # Paragraph properties (alignment, numbering, etc.)
    pPr = p.find("./w:pPr", NS)

    # Build buckets from original runs
    buckets = build_run_buckets(p)

    # If paragraph has no runs, create synthetic run
    if not buckets:
        synthetic = ET.Element(w_tag("r"))
        buckets = [RunBucket(run_template=synthetic, token_budget=999999)]

    # New paragraph child elements
    new_children: List[ET.Element] = []
    if pPr is not None:
        new_children.append(deepcopy(pPr))

    # Bucket pointer
    b_idx = 0               # which run bucket we are filling
    b_used = 0              # tokens used in current bucket
    current_text_parts: List[str] = []   # actual output text parts for current run

    def ensure_bucket():
        """
        Ensure we always have a bucket.
        If updated text is longer than original runs, reuse last run formatting.
        """
        nonlocal buckets, b_idx
        if b_idx >= len(buckets):
            last = buckets[-1].run_template
            buckets.append(RunBucket(run_template=last, token_budget=999999))

    def flush_run_bucket(force_advance=True):
        """
        Flush the current_text_parts into a run node with formatting of current bucket.

        force_advance=True:
          means close this run and move to next bucket.
        """
        nonlocal b_idx, b_used, current_text_parts
        if b_idx >= len(buckets):
            return

        tmpl = buckets[b_idx].run_template
        text = "".join(current_text_parts)  # ✅ exact spacing preserved
        new_children.append(clone_run_with_text(tmpl, text))

        current_text_parts = []
        b_used = 0

        if force_advance:
            b_idx += 1

    def flush_before_change_node():
        """
        Before inserting <w:ins>/<w:del>, we must close the previous run.

        This ensures inserted/deleted text does not mix into same normal run.
        """
        nonlocal b_idx
        if current_text_parts:
            flush_run_bucket(True)
        else:
            # even empty bucket gets "closed" to preserve boundaries
            if b_idx < len(buckets):
                flush_run_bucket(True)

    def add_token(tok: str):
        """
        Add a token to the current run.
        When bucket reaches its token budget, flush and move to next bucket.
        """
        nonlocal b_used
        ensure_bucket()
        current_text_parts.append(tok)
        b_used += 1

        # If current run reached its original token budget -> flush
        if buckets[b_idx].token_budget > 0 and b_used >= buckets[b_idx].token_budget:
            flush_run_bucket(True)

    # Walk updated plan sequentially
    for item in plan.items:
        if isinstance(item, KeepTok):
            add_token(item.token)

        elif isinstance(item, InsBlock):
            ins_text = (item.text or "").strip("\n")
            if ins_text == "":
                continue

            flush_before_change_node()
            ensure_bucket()

            tmpl = buckets[b_idx].run_template
            new_children.append(make_ins(tmpl, ins_text, next_id, author, date))
            next_id += 1

        elif isinstance(item, DelBlock):
            del_text = (item.text or "").strip("\n")
            if del_text == "":
                continue

            flush_before_change_node()
            ensure_bucket()

            tmpl = buckets[b_idx].run_template
            new_children.append(make_del(tmpl, del_text, next_id, author, date))
            next_id += 1

    # Flush remaining normal text
    if current_text_parts:
        flush_run_bucket(True)

    # Keep remaining original runs as empty text (preserves structure)
    while b_idx < len(buckets):
        tmpl = buckets[b_idx].run_template
        new_children.append(clone_run_with_text(tmpl, ""))
        b_idx += 1

    # Replace paragraph children
    for ch in list(p):
        p.remove(ch)
    for ch in new_children:
        p.append(ch)

    return next_id


# ============================
# Apply whole document.xml
# ============================

def rebuild_document_xml_preserve_runs(
    document_xml_path: str,
    updated_txt_path: str,
    out_xml_path: str,
    author: str = "Tanmay",
):
    """
    Main driver:
      - parse updated.txt
      - load document.xml
      - for each paragraph that exists in updated.txt:
          rebuild with preserve-runs algorithm
      - write out new xml
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

        plan = parse_updated_paragraph(edits[i])
        next_id = rebuild_paragraph_preserve_runs(p, plan, author, next_id)
        changed += 1

    tree.write(out_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Created new XML: {out_xml_path}")
    print(f"✅ Updated {changed} paragraphs.")


if __name__ == "__main__":
    rebuild_document_xml_preserve_runs(
        document_xml_path="word/document.xml",
        updated_txt_path="updated.txt",
        out_xml_path="word/document_updated.xml",
        author="Tanmay",
    )
