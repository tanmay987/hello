"""
build_documentxml_trackchanges_preserve_runs.py

Goal
----
Create a NEW WordprocessingML document.xml (output) using:
  - updated.txt as ground truth for words/content and explicit <INS:>/<DEL:> blocks
  - document.xml as ground truth for formatting and run structure (w:r boundaries)

Key policies
------------
1) Preserve formatting: keep the SAME run structure from document.xml as much as possible.
2) updated.txt is truth: paragraph content comes from updated.txt.
3) Track changes:
   - <INS: ...> -> <w:ins> with ONE run containing the entire INS text
   - <DEL: ...> -> <w:del> with ONE run containing the entire DEL text
4) Matching strategy (two-phase):
   - Phase 1: keep matching prefix tokens doc==updated outside markers
   - Phase 2: after first mismatch, STOP matching; fill runs sequentially with updated content.
     No forward searching for matches. No token lookahead.
5) FIXED: token joining now preserves spaces between words (prevents "allnon" collisions).
"""

import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime, timezone
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict, Union

# ============================
# Namespace
# ============================
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ET.register_namespace("w", NS["w"])


def w_tag(tag: str) -> str:
    """Return fully qualified WordprocessingML tag."""
    return f"{{{NS['w']}}}{tag}"


def now_w3c() -> str:
    """UTC timestamp in Word track-changes format."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ============================
# updated.txt parsing
# ============================
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")
MARK_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)

# Token = word OR punctuation. Spaces are NOT tokens.
TOKEN_RE = re.compile(r"\w+|[^\w\s]", re.UNICODE)


def parse_updated_txt(path: str) -> Dict[int, str]:
    """
    Parse updated.txt lines:
      [P1] Some text <INS: new words> and <DEL: old words>

    Returns:
      dict { paragraph_index (1-based) -> updated paragraph string }
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
    Tokenize text into word + punctuation tokens.
    Spaces are removed at token level.

    Example:
      "school," -> ["school", ","]
      "in-car"  -> ["in", "-", "car"]
    """
    return [m.group(0) for m in TOKEN_RE.finditer(s)]


# ============================
# Updated paragraph plan
# ============================
@dataclass
class KeepToken:
    """A normal (visible) token from updated.txt outside INS/DEL."""
    token: str


@dataclass
class InsBlock:
    """Atomic insertion block. Keep as one <w:ins>."""
    text: str


@dataclass
class DelBlock:
    """Atomic deletion block. Keep as one <w:del>."""
    text: str


UpdatedItem = Union[KeepToken, InsBlock, DelBlock]


@dataclass
class UpdatedPlan:
    """
    Parsed plan for one updated paragraph.

    items: sequence of KeepToken / InsBlock / DelBlock in order.
    keep_tokens_only: list of KeepToken tokens only (for matching prefix).
    """
    items: List[UpdatedItem]
    keep_tokens_only: List[str]


def parse_updated_paragraph(updated_marked: str) -> UpdatedPlan:
    """
    Parse one updated paragraph into a sequence:
      KeepToken(token) for outside markers
      InsBlock(text) for <INS:...>
      DelBlock(text) for <DEL:...>
    """
    items: List[UpdatedItem] = []
    keep_only: List[str] = []

    pos = 0
    for m in MARK_RE.finditer(updated_marked):
        outside = updated_marked[pos:m.start()]
        for t in tokenize_words_punct(outside):
            items.append(KeepToken(t))
            keep_only.append(t)

        kind = m.group(1)
        inner = m.group(2)

        if kind == "INS":
            items.append(InsBlock(inner))
        else:
            items.append(DelBlock(inner))

        pos = m.end()

    tail = updated_marked[pos:]
    for t in tokenize_words_punct(tail):
        items.append(KeepToken(t))
        keep_only.append(t)

    return UpdatedPlan(items=items, keep_tokens_only=keep_only)


# ============================
# document.xml helpers
# ============================
def extract_run_text(run: ET.Element) -> str:
    """Concatenate all w:t text inside a run."""
    parts = []
    for t in run.findall(".//w:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def clone_run_with_text(run_template: ET.Element, text: str) -> ET.Element:
    """
    Clone formatting from run_template and set its text.
    Uses xml:space="preserve" to avoid collapsing spaces.
    """
    r = deepcopy(run_template)

    # Remove direct w:t children
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    if text != "":
        t = ET.Element(w_tag("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        r.append(t)

    return r


def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str) -> ET.Element:
    """Create <w:ins> with ONE run containing entire inserted text."""
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)
    ins.append(clone_run_with_text(run_template, text))
    return ins


def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str) -> ET.Element:
    """Create <w:del> with ONE run containing entire deleted text."""
    dele = ET.Element(w_tag("del"))
    dele.set(w_tag("id"), str(del_id))
    dele.set(w_tag("author"), author)
    dele.set(w_tag("date"), date)

    r = deepcopy(run_template)

    # Remove direct <w:t>
    for node in list(r):
        if node.tag == w_tag("t"):
            r.remove(node)

    dt = ET.Element(w_tag("delText"))
    dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    dt.text = text
    r.append(dt)

    dele.append(r)
    return dele


def smart_join_tokens(tokens: List[str]) -> str:
    """
    Join tokens into readable text WITHOUT removing spaces between words.

    Conservative spacing rules:
      - default: insert ONE space between tokens
      - NO space BEFORE: , . ; : ? ! ) ] }
      - NO space AFTER:  ( [ {
      - also no space around hyphen '-' for "in-car" style joins
    """
    if not tokens:
        return ""

    NO_SPACE_BEFORE = {",", ".", ";", ":", "?", "!", ")", "]", "}", "%"}
    NO_SPACE_AFTER = {"(", "[", "{"}

    out = [tokens[0]]

    for prev, cur in zip(tokens, tokens[1:]):
        # Hyphen joins: in-car, 10-12
        if prev == "-" or cur == "-":
            out.append(cur)
            continue

        # no space after opening brackets
        if prev in NO_SPACE_AFTER:
            out.append(cur)
            continue

        # no space before closing punctuation
        if cur in NO_SPACE_BEFORE:
            out.append(cur)
            continue

        # default: add a space
        out.append(" " + cur)

    return "".join(out)


# ============================
# Core: preserve runs, fill content from updated.txt
# ============================
@dataclass
class RunBucket:
    """
    One original run bucket from document.xml.

    token_budget is how many doc tokens were originally in that run.
    We fill updated tokens into the same run pattern.
    """
    run_template: ET.Element
    token_budget: int


def build_run_buckets(p: ET.Element) -> List[RunBucket]:
    """
    Build run buckets from paragraph direct runs:
      - each bucket retains formatting template
      - token_budget is the original token count in that run
    """
    buckets: List[RunBucket] = []
    for r in p.findall("./w:r", NS):
        txt = extract_run_text(r)
        toks = tokenize_words_punct(txt)
        buckets.append(RunBucket(run_template=r, token_budget=len(toks)))
    return buckets


def extract_doc_prefix_tokens(p: ET.Element) -> List[str]:
    """
    Extract paragraph tokens from doc direct runs in order.
    Used only for matching prefix.
    """
    toks: List[str] = []
    for r in p.findall("./w:r", NS):
        toks.extend(tokenize_words_punct(extract_run_text(r)))
    return toks


def rebuild_paragraph_preserve_runs(
    p: ET.Element,
    plan: UpdatedPlan,
    author: str,
    next_id: int,
) -> int:
    """
    Rebuild ONE paragraph with updated.txt as truth but preserve run formatting:

    - Uses existing runs as buckets, filling tokens based on each run's original token budget.
    - Inserts <w:ins>/<w:del> as separate nodes, closing current run first.
    - Does not create run-per-word.
    """
    date = now_w3c()

    # Keep paragraph properties (pPr) as-is
    pPr = p.find("./w:pPr", NS)

    # Build run buckets (direct children runs only)
    buckets = build_run_buckets(p)

    # If no direct runs exist, create a synthetic bucket
    if not buckets:
        synthetic_run = ET.Element(w_tag("r"))
        buckets = [RunBucket(run_template=synthetic_run, token_budget=999999)]

    # Prefix match length (not heavily used; kept for semantics)
    doc_prefix_tokens = extract_doc_prefix_tokens(p)
    upd_keep = plan.keep_tokens_only

    prefix_len = 0
    while prefix_len < len(doc_prefix_tokens) and prefix_len < len(upd_keep):
        if doc_prefix_tokens[prefix_len] == upd_keep[prefix_len]:
            prefix_len += 1
        else:
            break

    # Output paragraph children
    new_children: List[ET.Element] = []
    if pPr is not None:
        new_children.append(deepcopy(pPr))

    # Bucket pointer
    b_idx = 0
    b_used = 0
    current_tokens: List[str] = []

    def ensure_bucket():
        nonlocal buckets, b_idx
        if b_idx >= len(buckets):
            # reuse last run formatting if updated is longer
            last_tmpl = buckets[-1].run_template
            buckets.append(RunBucket(run_template=last_tmpl, token_budget=999999))

    def flush_run_bucket(force_advance: bool = True):
        """
        Flush current_tokens into a run using the current bucket template.

        force_advance=True means "close this run" and move to next run.
        We use force_advance when we must insert INS/DEL (close previous run).
        """
        nonlocal b_idx, b_used, current_tokens
        if b_idx >= len(buckets):
            return
        tmpl = buckets[b_idx].run_template
        text = smart_join_tokens(current_tokens)
        new_children.append(clone_run_with_text(tmpl, text))
        current_tokens = []
        b_used = 0
        if force_advance:
            b_idx += 1

    def add_token_to_current_run(tok: str):
        nonlocal b_used
        ensure_bucket()
        current_tokens.append(tok)
        b_used += 1
        # If we filled the bucket budget, flush and advance
        if buckets[b_idx].token_budget > 0 and b_used >= buckets[b_idx].token_budget:
            flush_run_bucket(force_advance=True)

    def flush_before_change_node():
        """Close current run before inserting a <w:ins>/<w:del> node."""
        nonlocal current_tokens
        if current_tokens:
            flush_run_bucket(force_advance=True)
        else:
            # Even if empty, we "close" current run and move forward
            if b_idx < len(buckets):
                flush_run_bucket(force_advance=True)

    # Emit updated items sequentially
    for item in plan.items:
        if isinstance(item, KeepToken):
            add_token_to_current_run(item.token)

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

    # Flush remaining tokens to current bucket (close it)
    if current_tokens:
        flush_run_bucket(force_advance=True)

    # Preserve remaining original runs as empty
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
# Apply to whole document.xml
# ============================
def rebuild_document_xml_preserve_runs(
    document_xml_path: str,
    updated_txt_path: str,
    out_xml_path: str,
    author: str = "Tanmay",
):
    """
    Build a NEW document xml:
      - updated.txt provides content + INS/DEL
      - document.xml provides formatting + run boundaries
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
