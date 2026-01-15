import re                                                  # Regex parsing
import xml.etree.ElementTree as ET                          # XML read/write
from copy import deepcopy                                   # Preserve formatting by cloning runs
from datetime import datetime, timezone                     # Track change timestamp

# ----------------------------
# WordprocessingML Namespace
# ----------------------------
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}  # WordprocessingML namespace
ET.register_namespace("w", NS["w"])                         # keep "w:" prefix in output XML

def w_tag(tag: str) -> str:                                 # convert 'p' -> '{namespace}p'
    return f"{{{NS['w']}}}{tag}"

def now_w3c() -> str:                                       # Word timestamps use W3C UTC format
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

# ----------------------------
# Parse updated.txt
# ----------------------------
LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")                 # parse lines like [P12] ...

TOKEN_RE = re.compile(r"<(INS|DEL):(.*?)>", re.DOTALL)       # parse tokens like <INS:...> <DEL:...>

def parse_updated_txt(updated_txt_path: str) -> dict[int, str]:
    """
    Reads updated.txt and returns dict:
        { paragraph_id (1-based): content_string }
    """
    edits = {}                                               # paragraph id -> updated marked text
    with open(updated_txt_path, "r", encoding="utf-8") as f:  # open updated.txt
        for line in f:                                       # read line by line
            line = line.rstrip("\n")                         # strip newline
            if not line.strip():                             # skip empty lines
                continue
            m = LINE_RE.match(line)                          # match paragraph prefix
            if not m:                                        # if format doesn't match, ignore
                continue
            pid = int(m.group(1))                            # paragraph number
            content = m.group(2)                             # content after [P..]
            edits[pid] = content                             # store
    return edits                                             # return mapping


def tokenize_marked_text(s: str):
    """
    Converts:
        'Hello <DEL:world><INS:Tanmay>'
    into:
        [('TEXT','Hello '), ('DEL','world'), ('INS','Tanmay')]
    """
    tokens = []                                              # output token list
    pos = 0                                                  # scanning pointer
    for m in TOKEN_RE.finditer(s):                           # iterate through <INS> <DEL> tokens
        if m.start() > pos:                                  # if plain text before token
            tokens.append(("TEXT", s[pos:m.start()]))        # add text chunk
        kind = m.group(1)                                    # INS or DEL
        content = m.group(2)                                 # inner text
        tokens.append((kind, content))                       # add token
        pos = m.end()                                        # move pointer
    if pos < len(s):                                         # trailing text after last token
        tokens.append(("TEXT", s[pos:]))                     # add trailing text
    return tokens                                            # return tokens


# ----------------------------
# Extract original paragraph text + run map
# ----------------------------
def paragraph_plain_text(p: ET.Element) -> str:
    """Concatenate visible text in paragraph from all <w:t> nodes."""
    parts = []                                               # collect fragments
    for t in p.findall(".//w:t", NS):                        # all visible text nodes
        if t.text:                                           # skip if None
            parts.append(t.text)                             # append text
    return "".join(parts)                                    # return paragraph text


def build_run_map(p: ET.Element):
    """
    Maps each <w:r> run's text to global offsets within paragraph text.
    Returns:
      full_text: str
      mapping: list of {r, text, start, end}
    """
    full_text = ""                                           # combined paragraph text
    mapping = []                                             # list of run entries

    for r in p.findall("./w:r", NS):                         # direct runs only
        texts = []                                           # gather text fragments inside the run
        for t in r.findall("./w:t", NS):                     # normal text nodes in run
            if t.text:                                       # only non-empty
                texts.append(t.text)                         # add fragment
        run_text = "".join(texts)                            # run text

        if run_text:                                         # only include runs that contain text
            start = len(full_text)                           # start offset
            full_text += run_text                            # append run text
            end = len(full_text)                             # end offset
            mapping.append({"r": r, "text": run_text, "start": start, "end": end})  # store mapping

    return full_text, mapping                                # return overall text and mapping


# ----------------------------
# Run splitting (preserves formatting)
# ----------------------------
def split_run_at(run: ET.Element, offset: int):
    """
    Split a run into two runs at 'offset' in its text.
    Returns: (left_run, right_run)
    Both runs preserve formatting (clone original run).
    """
    t_nodes = run.findall("./w:t", NS)                       # all <w:t> nodes
    combined = "".join([(t.text or "") for t in t_nodes])    # combined run text

    left_text = combined[:offset]                            # left part
    right_text = combined[offset:]                           # right part

    left = deepcopy(run)                                     # clone run for left
    right = deepcopy(run)                                    # clone run for right

    for node in list(left):                                  # remove old text nodes in left
        if node.tag == w_tag("t"):
            left.remove(node)

    for node in list(right):                                 # remove old text nodes in right
        if node.tag == w_tag("t"):
            right.remove(node)

    def add_wt(run_elem: ET.Element, text: str):             # helper to add <w:t>
        if text == "":                                       # skip empty
            return
        t = ET.Element(w_tag("t"))                           # create <w:t>
        if text.startswith(" ") or text.endswith(" "):       # preserve spaces
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text                                        # assign content
        run_elem.append(t)                                   # append to run

    add_wt(left, left_text)                                  # put left text
    add_wt(right, right_text)                                # put right text

    return left, right                                       # return both halves


# ----------------------------
# Build track change XML nodes
# ----------------------------
def make_del(run_template: ET.Element, text: str, del_id: int, author: str, date: str) -> ET.Element:
    dele = ET.Element(w_tag("del"))                          # create <w:del>
    dele.set(w_tag("id"), str(del_id))                       # w:id
    dele.set(w_tag("author"), author)                        # w:author
    dele.set(w_tag("date"), date)                            # w:date

    r = deepcopy(run_template)                               # clone formatting
    for node in list(r):                                     # remove normal <w:t>
        if node.tag == w_tag("t"):
            r.remove(node)

    dt = ET.Element(w_tag("delText"))                        # create <w:delText>
    if text.startswith(" ") or text.endswith(" "):           # preserve spaces
        dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    dt.text = text                                           # deleted content
    r.append(dt)                                             # add to run
    dele.append(r)                                           # add run to del wrapper
    return dele                                              # return <w:del>


def make_ins(run_template: ET.Element, text: str, ins_id: int, author: str, date: str) -> ET.Element:
    ins = ET.Element(w_tag("ins"))                           # create <w:ins>
    ins.set(w_tag("id"), str(ins_id))                        # w:id
    ins.set(w_tag("author"), author)                         # w:author
    ins.set(w_tag("date"), date)                             # w:date

    r = deepcopy(run_template)                               # clone formatting
    for node in list(r):                                     # remove normal <w:t>
        if node.tag == w_tag("t"):
            r.remove(node)

    t = ET.Element(w_tag("t"))                               # new inserted <w:t>
    if text.startswith(" ") or text.endswith(" "):           # preserve spaces
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text                                            # inserted content
    r.append(t)                                              # add <w:t> to run
    ins.append(r)                                            # add run to ins wrapper
    return ins                                               # return <w:ins>


# ----------------------------
# Apply tokens to paragraph
# ----------------------------
def apply_tokens_to_paragraph(p: ET.Element, tokens, author: str, next_id: int):
    """
    Align tokens to original paragraph text and inject track changes in-place.
    This assumes TEXT tokens appear in original paragraph in order.
    """
    original_text, mapping = build_run_map(p)                # build mapping over runs
    cursor = 0                                               # current char offset in original_text
    date = now_w3c()                                         # single timestamp per paragraph change

    parent_children = list(p)                                # paragraph children list (runs etc.)

    def find_child_index(elem):                              # locate element index in p children
        for i, ch in enumerate(parent_children):
            if ch is elem:
                return i
        return None

    # choose formatting template: first run with text, else create dummy run
    template_run = mapping[0]["r"] if mapping else ET.Element(w_tag("r"))

    new_children = []                                        # rebuild paragraph child list
    child_i = 0                                              # pointer over old paragraph children

    def consume_children_until_run(target_run):              # copy children until we reach target_run
        nonlocal child_i
        while child_i < len(parent_children):
            ch = parent_children[child_i]
            if ch is target_run:
                break
            new_children.append(ch)                          # keep untouched non-run children
            child_i += 1

    for kind, content in tokens:                             # iterate tokens from updated text
        if kind == "TEXT":                                   # unchanged text
            if content == "":
                continue

            # Find this unchanged content in the original text at/after cursor
            pos = original_text.find(content, cursor)        # locate content starting at cursor
            if pos < 0:
                raise ValueError(f"Cannot align TEXT token '{content}' in original paragraph '{original_text}'")

            # We copy original paragraph children up to the runs that represent this part.
            # For simplicity we will NOT split/trim runs for TEXT - we keep them as-is.
            cursor = pos + len(content)                      # advance cursor beyond matched text

        elif kind == "DEL":                                  # deletion relative to original
            if content == "":
                continue

            # Deletion must exist in original text at/after cursor
            pos = original_text.find(content, cursor)        # find deleted substring
            if pos < 0:
                raise ValueError(f"Cannot locate DEL text '{content}' in original paragraph '{original_text}'")

            del_start = pos                                  # start offset
            del_end = pos + len(content)                     # end offset
            cursor = del_end                                 # advance cursor in original

            # Find first run overlapping del_start
            affected = []
            for m in mapping:
                if m["end"] <= del_start:
                    continue
                if m["start"] >= del_end:
                    continue
                affected.append(m)

            if not affected:
                continue

            first = affected[0]
            last = affected[-1]
            first_run = first["r"]
            last_run = last["r"]

            # Copy children up to first affected run
            consume_children_until_run(first_run)            # keep earlier children unchanged

            first_local_start = del_start - first["start"]   # local offset in first run
            last_local_end = del_end - last["start"]         # local offset in last run

            # Split first affected run at start
            left_run, first_right = split_run_at(first_run, first_local_start)

            # Split last affected run at end
            if first_run is last_run:
                mid_run, right_run = split_run_at(first_right, last_local_end - first_local_start)
            else:
                mid_run, right_run = split_run_at(last_run, last_local_end)

            # Append left unchanged piece
            if left_run.findall("./w:t", NS):
                new_children.append(left_run)

            # Insert <w:del> node using first run formatting
            del_node = make_del(first_run, content, next_id, author, date)
            next_id += 1
            new_children.append(del_node)

            # Skip over all children until after last affected run
            # Move child_i pointer beyond last_run
            while child_i < len(parent_children):
                ch = parent_children[child_i]
                child_i += 1
                if ch is last_run:
                    break

            # Append right unchanged piece
            if right_run.findall("./w:t", NS):
                new_children.append(right_run)

        elif kind == "INS":                                  # insertion relative to original
            if content == "":
                continue

            # Insertions don't consume original characters.
            # We just insert at current cursor position.
            ins_node = make_ins(template_run, content, next_id, author, date)
            next_id += 1
            new_children.append(ins_node)

    # After processing tokens, append remaining untouched children
    while child_i < len(parent_children):                     # copy leftover children
        new_children.append(parent_children[child_i])
        child_i += 1

    # Replace paragraph children with new children list
    for ch in list(p):
        p.remove(ch)
    for ch in new_children:
        p.append(ch)

    return next_id


# ----------------------------
# Main apply function
# ----------------------------
def apply_changes(document_xml_path: str, updated_txt_path: str, author: str = "Tanmay"):
    edits = parse_updated_txt(updated_txt_path)               # read updated.txt mapping

    tree = ET.parse(document_xml_path)                        # parse document.xml
    root = tree.getroot()                                     # root node
    paragraphs = root.findall(".//w:p", NS)                   # list all paragraphs in doc

    next_id = 1                                               # id counter for tracked changes
    changed = 0                                               # how many paragraphs changed

    for i, p in enumerate(paragraphs, start=1):               # 1-based paragraph numbers
        if i not in edits:                                    # if no update line for this paragraph
            continue                                          # skip

        updated = edits[i]                                    # updated paragraph content
        tokens = tokenize_marked_text(updated)                # tokenize updated content

        # Only apply if paragraph contains any INS/DEL tokens
        if not any(k in ("INS", "DEL") for k, _ in tokens):
            continue

        next_id = apply_tokens_to_paragraph(p, tokens, author, next_id)  # inject changes
        changed += 1                                          # count

    tree.write(document_xml_path, encoding="utf-8", xml_declaration=True)  # save updated XML
    print(f"✅ Updated {changed} paragraphs with tracked changes.")


# ---- USAGE (after you have extracted docx zip) ----
# apply_changes("word/document.xml", "updated.txt", author="Tanmay")
