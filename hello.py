import re

LINE_RE = re.compile(r"^\[P(\d+)\]\s?(.*)$")                    # matches "[P12] rest of paragraph text"

def parse_updated_txt(updated_txt_path: str) -> dict[int, str]:
    edits = {}                                                  # paragraph_id -> updated marked paragraph

    with open(updated_txt_path, "r", encoding="utf-8") as f:    # open updated.txt
        for line in f:                                          # read line-by-line
            line = line.rstrip("\n")                            # remove newline at end
            if not line.strip():                                # skip empty lines
                continue

            m = LINE_RE.match(line)                             # extract paragraph id
            if not m:                                           # ignore lines not in format
                continue

            pid = int(m.group(1))                               # paragraph number
            content = m.group(2)                                # the paragraph content after [P...]
            edits[pid] = content                                # store mapping

    return edits                                                 # return dict

# ---- Usage ----
if __name__ == "__main__":
    edits = parse_updated_txt("updated.txt")                    # parse updated.txt
    print(edits)                                                # print dictionary
