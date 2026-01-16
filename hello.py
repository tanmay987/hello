"""
build_documentxml_trackchanges_preserve_runs_spaces.py

updated.txt is ground truth INCLUDING SPACES.
document.xml is ground truth for formatting and run boundaries.

This script preserves spaces by keeping whitespace as real tokens.
"""

import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime, timezone
from dataclasses import dataclass
from typing import List, Dict, Union

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

# ✅ SPACE-PRESERVING tokenizer
# token could be whitespace OR word OR punctuation
TOKEN_RE = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)


def parse_updated_txt(path: str) -> Dict[int, str]:
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


def tokenize_with_spaces(s: str) -> List[str]:
    """
    Tokenize into:
      - whitespace tokens
      - word tokens
      - punctuation tokens
    No information is lost.
    """
    return [m.group(0) for m in TOKEN_RE.finditer(s)]


# ============================
# Updated paragraph plan
# ============================
@dataclass
class KeepTok:
    token: str  # can be space/word/punct

@dataclass
class InsBlock:
    text: str

@dataclass
class DelBlock:
    text: str

UpdatedItem = Union[KeepTok, InsBlock, DelBlock]

@dataclass
class UpdatedPlan:
    items: List[UpdatedItem]


def parse_updated_paragraph(updated_marked: str) -> UpdatedPlan:
    items: List[UpdatedItem] = []
    pos = 0

    for m in MARK_RE.finditer(updated_marked):
        outside = updated_marked[pos:m.start()]
        for t in tokenize_with_spaces(outside):
            items.append(KeepTok(t))

        kind = m.group(1)
        inner = m.group(2)

        if kind == "INS":
            items.append(InsBlock(inner))
        else:
            items.append(DelBlock(inner))

        pos = m.end()

    tail = updated_marked[pos:]
    for t in tokenize_with_spaces(tail):
        items.append(KeepTok(t))

    return UpdatedPlan(items=items)


# ============================
# document.xml helpers
# ============================
def extract_run_text(run: ET.Element) -> str:
    parts = []
    for t in run.findall(".//w:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def clone_run_with_text(run_template: ET.Element, text: str) -> ET.Element:
    """
    Clone run formatting and set text exactly.
    IMPORTANT: xml:space="preserve" ensures spaces remain.
    """
    r = deepcopy(run_template)

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
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)
    ins.append(clone_run_with_text(run_template, text))
    return ins


def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str) -> ET.Element:
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
# Preserve runs using "token budget"
# ============================
@dataclass
class RunBucket:
    run_template: ET.Element
    token_budget: int


def build_run_buckets(p: ET.Element) -> List[RunBucket]:
    """
    Use direct paragraph runs as buckets.
    token_budget includes whitespace tokens so spacing shape matches doc structure.
    """
    buckets: List[RunBucket] = []
    for r in p.findall("./w:r", NS):
        txt = extract_run_text(r)
        toks = tokenize_with_spaces(txt)
        buckets.append(RunBucket(run_template=r, token_budget=len(toks)))
    return buckets


def rebuild_paragraph_preserve_runs(p: ET.Element, plan: UpdatedPlan, author: str, next_id: int) -> int:
    """
    Fill the SAME run structure using updated tokens (including whitespace),
    and insert INS/DEL as tracked changes.
    """
    date = now_w3c()

    pPr = p.find("./w:pPr", NS)

    buckets = build_run_buckets(p)
    if not buckets:
        synthetic = ET.Element(w_tag("r"))
        buckets = [RunBucket(run_template=synthetic, token_budget=999999)]

    new_children: List[ET.Element] = []
    if pPr is not None:
        new_children.append(deepcopy(pPr))

    b_idx = 0
    b_used = 0
    current_text_parts: List[str] = []

    def ensure_bucket():
        nonlocal buckets, b_idx
        if b_idx >= len(buckets):
            last = buckets[-1].run_template
            buckets.append(RunBucket(run_template=last, token_budget=999999))

    def flush_run_bucket(force_advance=True):
        nonlocal b_idx, b_used, current_text_parts
        if b_idx >= len(buckets):
            return
        tmpl = buckets[b_idx].run_template
        text = "".join(current_text_parts)  # ✅ no join guesswork
        new_children.append(clone_run_with_text(tmpl, text))
        current_text_parts = []
        b_used = 0
        if force_advance:
            b_idx += 1

    def flush_before_change_node():
        nonlocal b_idx
        if current_text_parts:
            flush_run_bucket(True)
        else:
            # close even empty run bucket to preserve boundaries
            if b_idx < len(buckets):
                flush_run_bucket(True)

    def add_token(tok: str):
        nonlocal b_used
        ensure_bucket()
        current_text_parts.append(tok)
        b_used += 1
        if buckets[b_idx].token_budget > 0 and b_used >= buckets[b_idx].token_budget:
            flush_run_bucket(True)

    for item in plan.items:
        if isinstance(item, KeepTok):
            add_token(item.token)

        elif isinstance(item, InsBlock):
            ins_text = item.text
            if ins_text is None:
                continue
            # preserve INS text exactly (including spaces inside)
            ins_text = ins_text.strip("\n")
            if ins_text == "":
                continue

            flush_before_change_node()
            ensure_bucket()
            tmpl = buckets[b_idx].run_template
            new_children.append(make_ins(tmpl, ins_text, next_id, author, date))
            next_id += 1

        elif isinstance(item, DelBlock):
            del_text = item.text
            if del_text is None:
                continue
            del_text = del_text.strip("\n")
            if del_text == "":
                continue

            flush_before_change_node()
            ensure_bucket()
            tmpl = buckets[b_idx].run_template
            new_children.append(make_del(tmpl, del_text, next_id, author, date))
            next_id += 1

    if current_text_parts:
        flush_run_bucket(True)

    # preserve remaining runs as empty
    while b_idx < len(buckets):
        tmpl = buckets[b_idx].run_template
        new_children.append(clone_run_with_text(tmpl, ""))
        b_idx += 1

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
