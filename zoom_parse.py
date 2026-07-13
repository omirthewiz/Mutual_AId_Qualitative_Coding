import re, sys
from pathlib import Path
# turns a zoom transcript-panel copy/paste into a clean "Name: text" format
# that clean.py can read. just fixes formatting - does NOT anonymize.
# zoom format: the speaker name shows up twice in a row, then an optional
# timestamp, then the spoken lines, until the next double-name.
# how to run: python3 zoom_parse.py ws3_raw.txt -o workshop3.vtt
# then anonymize + clean with: python3 clean.py workshop3.vtt -n "Real Name" -a "P1"

is_time = re.compile(r'^\d{1,2}:\d{2}:\d{2}$|^\d{1,2}:\d{2}$')

def parse(path):
    lines = Path(path).read_text(encoding="utf-8").replace("\r\n","\n").split("\n")
    turns = []
    i = 0
    n = len(lines)
    speaker = None
    time = ""
    text = []
    def flush():
        if speaker is not None and text:
            turns.append((speaker, time, " ".join(text).strip()))
    while i < n:
        line = lines[i].strip()
        nxt = lines[i+1].strip() if i+1 < n else ""
        # a speaker label = a non-empty line repeated on the next line,
        # and it's short (a name, not a sentence)
        if line and line == nxt and len(line) < 40 and not is_time.match(line):
            flush()
            speaker = line
            time = ""
            text = []
            i += 2
            # optional timestamp right after the name
            if i < n and is_time.match(lines[i].strip()):
                time = lines[i].strip()
                i += 1
            continue
        if line:
            text.append(line)
        i += 1
    flush()
    return turns

def main():
    if len(sys.argv) < 2:
        sys.exit("Usage: python3 zoom_parse.py raw.txt -o workshop3.vtt")
    inp = sys.argv[1]
    out = "workshop3.vtt"
    if "-o" in sys.argv:
        out = sys.argv[sys.argv.index("-o")+1]
    turns = parse(inp)
    if not turns:
        sys.exit("Nothing parsed - check the paste.")
    # merge consecutive same-speaker turns
    merged = []
    for speaker, time, text in turns:
        if merged and merged[-1][0] == speaker:
            merged[-1] = (speaker, merged[-1][1], merged[-1][2] + " " + text)
        else:
            merged.append((speaker, time, text))
    with open(out, "w", encoding="utf-8") as f:
        for speaker, time, text in merged:
            if time:
                t = time if time.count(":") == 2 else "00:" + time
                f.write(t + ".0\n")
            else:
                f.write("00:00:00.0\n")
            f.write(f"{speaker}: {text}\n\n")
    print(f"Wrote {len(merged)} turns to {out}")
    # list the speakers found, so you know who to anonymize with clean.py
    seen = []
    for s,_,_ in merged:
        if s not in seen: seen.append(s)
    print("Speakers found (anonymize these with clean.py -n/-a):")
    for s in seen: print("  ", s)

main()