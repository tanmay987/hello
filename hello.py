import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime, timezone

# ----------------------------
# WordprocessingML namespace
# ----------------------------
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
ET.register_namespace("w", NS["w"])


def w_tag(tag: str) -> str:
    """Return a fully-qualified WordprocessingML tag name for ElementTree.

    Word XML uses a namespace URI. ElementTree represents namespaced tags as:
        "{namespace-uri}tag"
    This helper converts a short tag like "p" or "r" into the namespaced form.

    Args:
        tag: The local tag name (e.g., "p", "r", "t", "ins", "del").

    Returns:
        A namespaced tag string usable with ElementTree (e.g., "{...}p").
    """
    return f"{{{NS['w']}}}{tag}"


def now_w3c() -> str:
    """Return the current UTC time formatted for Word track changes (W3C format).

    Word track-changes metadata expects a timestamp such as:
        2026-01-15T10:30:00Z

    Returns:
        A UTC timestamp string in the format "%Y-%m-%dT%H:%M:%SZ".
    """
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


# ----------------------------
# Updated.txt parsing
# ----------------------------
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")
TOKEN_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)


def parse_updated_txt(updated_txt_path: str) -> dict[int, str]:
    """Parse updated.txt into a mapping of paragraph-id -> marked content.

    Expected updated.txt format:
        [P1] Hello <DEL:world><INS:Tanmay>
        [P2] This is unchanged.
        [P3] I like <INS:Python>

    Args:
        updated_txt_path: Path to the updated.txt file.

    Returns:
        Dictionary mapping 1-based paragraph number to the full marked line content.
    """
    edits: dict[int, str] = {}
    with open(updated_txt_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line.strip():
                continue
            m = LINE_RE.match(line)
            if not m:
                continue
            pid = int(m.group(1))
            edits[pid] = m.group(2)
    return edits


def tokenize_marked_text(s: str):
    """Tokenize a marked paragraph string into ordered (kind, text) tokens.

    Converts a string like:
        "Hello <DEL:world><INS:Tanmay> !!"
    into:
        [("TEXT","Hello "), ("DEL","world"), ("INS","Tanmay"), ("TEXT"," !!")]

    Token kinds:
        - "TEXT": unchanged anchor text
        - "DEL": deleted text
        - "INS": inserted text

    Args:
        s: The marked paragraph text.

    Returns:
        List of (kind, content) tokens in order.
    """
    tokens = []
    pos = 0
    for m in TOKEN_RE.finditer(s):
        if m.start() > pos:
            tokens.append(("TEXT", s[pos:m.start()]))
        tokens.append((m.group(1), m.group(2)))
        pos = m.end()
    if pos < len(s):
        tokens.append(("TEXT", s[pos:]))
    return tokens


# ----------------------------
# Extract original paragraph text
# ----------------------------
def paragraph_plain_text(p: ET.Element) -> str:
    """Extract the visible plain text from a Word paragraph element.

    This concatenates text across all <w:t> nodes (normal text nodes)
    within the paragraph.

    Notes:
        - This does NOT include <w:delText> from prior tracked deletions.
        - It does not try to preserve line breaks; it simply concatenates.

    Args:
        p: An ElementTree element representing <w:p>.

    Returns:
        The concatenated paragraph text as a string.
    """
    parts = []
    for t in p.findall(".//w:t", NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


# ----------------------------
# Run map and splitting helpers
# ----------------------------
def build_run_map(p: ET.Element):
    """Build a run-to-text offset map for a paragraph.

    Word paragraphs contain multiple runs (<w:r>). Each run may contain
    one or more <w:t> nodes. This function concatenates run texts to form
    the full paragraph string and records each run's [start,end) offsets.

    Only direct runs (./w:r) are mapped (not those nested inside hyperlinks).

    Args:
        p: Paragraph element <w:p>.

    Returns:
        A tuple (full_text, mapping) where:
            - full_text: concatenated paragraph text
            - mapping: list of dict entries with keys:
                {"r": run_element, "text": run_text, "start": int, "end": int}
    """
    full_text = ""
    mapping = []
    for r in p.findall("./w:r", NS):
        texts = []
        for t in r.findall("./w:t", NS):
            if t.text:
                texts.append(t.text)
        run_text = "".join(texts)
        if run_text:
            start = len(full_text)
            full_text += run_text
            end = len(full_text)
            mapping.append({"r": r, "text": run_text, "start": start, "end": end})
    return full_text, mapping


def split_run_at(run: ET.Element, offset: int):
    """Split a <w:r> run into two runs at a character offset.

    This is used to surgically insert <w:ins> / <w:del> at exact character
    positions without destroying run formatting.

    The run is cloned (deepcopy) so its formatting (w:rPr) is preserved.

    Args:
        run: The run element <w:r> to split.
        offset: Character index within the run's combined text at which to split.

    Returns:
        A tuple (left_run, right_run) as cloned run elements containing
        the respective text halves. If one side has empty text, it will
        have no <w:t> node.
    """
    t_nodes = run.findall("./w:t", NS)
    combined = "".join([(t.text or "") for t in t_nodes])

    left_text = combined[:offset]
    right_text = combined[offset:]

    left = deepcopy(run)
    right = deepcopy(run)

    # remove existing <w:t> nodes from clones
    for node in list(left):
        if node.tag == w_tag("t"):
            left.remove(node)
    for node in list(right):
        if node.tag == w_tag("t"):
            right.remove(node)

    def add_wt(run_elem: ET.Element, text: str):
        """Add a <w:t> node with xml:space preservation if needed."""
        if text == "":
            return
        t = ET.Element(w_tag("t"))
        if text.startswith(" ") or text.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
        run_elem.append(t)

    add_wt(left, left_text)
    add_wt(right, right_text)
    return left, right


# ----------------------------
# Track changes node builders
# ----------------------------
def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str):
    """Create a <w:del> tracked deletion element with a formatted run.

    Word stores deleted text inside <w:delText> (NOT <w:t>).

    Args:
        run_template: A run whose formatting should be cloned.
        text: Deleted text content.
        del_id: Unique track change ID (integer).
        author: Track change author name.
        date: Track change timestamp string (W3C UTC format).

    Returns:
        An ElementTree element representing <w:del>...</w:del>.
    """
    dele = ET.Element(w_tag("del"))
    dele.set(w_tag("id"), str(del_id))
    dele.set(w_tag("author"), author)
    dele.set(w_tag("date"), date)

    r = deepcopy(run_template)

    # delete nodes must use <w:delText>, not <w:t>
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
    """Create a <w:ins> tracked insertion element with a formatted run.

    Args:
        run_template: A run whose formatting should be cloned.
        text: Inserted text content.
        ins_id: Unique track change ID (integer).
        author: Track change author name.
        date: Track change timestamp string (W3C UTC format).

    Returns:
        An ElementTree element representing <w:ins>...</w:ins>.
    """
    ins = ET.Element(w_tag("ins"))
    ins.set(w_tag("id"), str(ins_id))
    ins.set(w_tag("author"), author)
    ins.set(w_tag("date"), date)

    r = deepcopy(run_template)

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


# ----------------------------
# Core: apply multiple marker changes in a paragraph
# ----------------------------
def apply_markers_to_paragraph(p: ET.Element, marked_text: str, author: str, next_id: int):
    """Apply <INS:...> and <DEL:...> markers to a paragraph as Word track changes.

    The algorithm treats TEXT outside markers as "anchors" that must appear in the
    original paragraph (in order). It uses these anchors to locate exact positions
    where deletions/insertions occur.

    Multiple markers are supported in a single paragraph.

    Important assumptions:
        - Each <DEL:...> content exists in the original paragraph text after cursor.
        - Anchor TEXT segments exist in the original paragraph text after cursor.
        - Paragraph numbering in updated.txt matches document.xml's paragraph ordering.

    Args:
        p: Paragraph element <w:p> to modify.
        marked_text: The marked paragraph string containing <INS:..> / <DEL:..>.
        author: Track-change author name.
        next_id: Next available track-change ID (will increment as changes are added).

    Returns:
        Updated next_id after inserting track-change nodes.
    """
    original_text, run_map = build_run_map(p)
    tokens = tokenize_marked_text(marked_text)
    date = now_w3c()
    cursor = 0

    parent_children = list(p)
    new_children = []
    child_i = 0

    template_run_default = run_map[0]["r"] if run_map else ET.Element(w_tag("r"))

    def consume_children_until_run(target_run: ET.Element):
        """Copy paragraph children into new_children until hitting target_run."""
        nonlocal child_i
        while child_i < len(parent_children):
            ch = parent_children[child_i]
            if ch is target_run:
                break
            new_children.append(ch)
            child_i += 1

    def find_run_at_offset(off: int):
        """Return run-map entry covering offset 'off' in original_text, else None."""
        for m in run_map:
            if m["start"] <= off < m["end"]:
                return m
        return None

    def apply_deletion(del_text: str):
        """Inject a tracked deletion (<w:del>) for del_text found after cursor."""
        nonlocal cursor, next_id, child_i

        if del_text == "":
            return

        pos = original_text.find(del_text, cursor)
        if pos < 0:
            raise ValueError(
                f"DEL text not found.\nDEL='{del_text}'\nFrom cursor={cursor}\nOriginal='{original_text}'"
            )

        start = pos
        end = pos + len(del_text)
        cursor = end

        affected = []
        for m in run_map:
            if m["end"] <= start:
                continue
            if m["start"] >= end:
                continue
            affected.append(m)

        if not affected:
            return

        first = affected[0]
        last = affected[-1]
        first_run = first["r"]
        last_run = last["r"]

        consume_children_until_run(first_run)

        first_local = start - first["start"]
        last_local = end - last["start"]

        left_run, first_right = split_run_at(first_run, first_local)

        if first_run is last_run:
            _mid, right_run = split_run_at(first_right, last_local - first_local)
        else:
            _mid, right_run = split_run_at(last_run, last_local)

        if left_run.findall("./w:t", NS):
            new_children.append(left_run)

        del_node = make_del(first_run, del_text, next_id, author, date)
        next_id += 1
        new_children.append(del_node)

        while child_i < len(parent_children):
            ch = parent_children[child_i]
            child_i += 1
            if ch is last_run:
                break

        if right_run.findall("./w:t", NS):
            new_children.append(right_run)

    def apply_insertion(ins_text: str):
        """Inject a tracked insertion (<w:ins>) at the current cursor position."""
        nonlocal next_id
        if ins_text == "":
            return

        run_here = find_run_at_offset(cursor)
        template = run_here["r"] if run_here else template_run_default

        ins_node = make_ins(template, ins_text, next_id, author, date)
        next_id += 1
        new_children.append(ins_node)

    for kind, content in tokens:
        if kind == "TEXT":
            if content == "":
                continue

            pos = original_text.find(content, cursor)
            if pos < 0:
                raise ValueError(
                    f"TEXT anchor not found.\nTEXT='{content}'\nFrom cursor={cursor}\nOriginal='{original_text}'"
                )

            cursor = pos + len(content)

        elif kind == "DEL":
            apply_deletion(content)

        elif kind == "INS":
            apply_insertion(content)

    while child_i < len(parent_children):
        new_children.append(parent_children[child_i])
        child_i += 1

    for ch in list(p):
        p.remove(ch)
    for ch in new_children:
        p.append(ch)

    return next_id


# ----------------------------
# Main: apply across document.xml
# ----------------------------
def apply_track_changes(document_xml_path: str, updated_txt_path: str, author: str = "Tanmay"):
    """Apply tracked changes to a document.xml based on updated.txt marker annotations.

    This reads:
        - document_xml_path: "word/document.xml"
        - updated_txt_path: paragraph-indexed updated text with <INS:...>/<DEL:...>

    For each paragraph with markers, it injects <w:ins> and <w:del> nodes
    into the corresponding <w:p> element.

    Args:
        document_xml_path: Path to the Word document.xml file (already extracted from .docx).
        updated_txt_path: Path to updated.txt file with paragraph lines.
        author: Track changes author string for Word metadata.

    Side effects:
        - Overwrites document_xml_path with modified XML containing track changes.
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

        next_id = apply_markers_to_paragraph(p, marked, author, next_id)
        changed += 1

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)
    print(f"✅ Applied tracked changes to {changed} paragraphs.")


if __name__ == "__main__":
    """Run marker-based track changes injection.

    Expected:
        - You have extracted a .docx so that 'word/document.xml' exists on disk.
        - You have an 'updated.txt' file with [P#] lines and markers.

    Adjust the paths if needed.
    """
    apply_track_changes("word/document.xml", "updated.txt", author="Tanmay")
