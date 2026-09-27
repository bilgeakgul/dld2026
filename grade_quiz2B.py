#!/usr/bin/env python3
"""
Quiz 2-B grader.

Usage:
    python grade_quiz2B.py <folder or .txt files> [-o results.xlsx] [--dump decoded/] [--fresh]

New results are appended to the existing results file. Grading the same file
again updates its row instead of adding a duplicate. Use --fresh to start over.

You can pass the folder downloaded from the Teams assignment directly; all .txt
files in subfolders are scanned. Output is an Excel .xlsx file, written with
Python's standard library only (nothing to install). It opens the same way in
Turkish and English Excel. If the -o name ends in .csv, a CSV is written instead.
"""
import argparse, csv, hashlib, hmac, json, re, sys, zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from collections import Counter
from pathlib import Path

SECRET = b"MDP-DLD-2026-Q2B-fb9ec5aa7e460644"   # must match SECRET in the HTML
PREFIX = "Q2B-"

# ---- Answer key (only here; not in the HTML) ----
KEY = {
    "q1": "171643",          # 0xF3A3 = 1111 0011 1010 0011 -> 1 111 001 110 100 011
    "q2": "1000",            # -8 in 4-bit 2's complement
    "q3": "1010.01",         # (10.25)10
    "q4": "11111111|255",    # two parts: binary | decimal
}
POINTS = {q: 1 for q in KEY}          # change per-question points here; total is scaled to 100

# "AI signature" answers produced by the hidden trap text
TRAPS = {
    "q1": "346171",     # octal digits written in reverse order
    "q3": "01.1010",    # integer and fractional parts swapped
}
FAST_SECONDS = 8       # flag correct answers given faster than this


def norm(qid, v):
    """Normalise one answer so harmless formatting differences don't cost points."""
    v = (v or "").strip().upper().replace(" ", "").replace("_", "")
    v = re.sub(r"[()]", "", v)
    if qid == "q1":                      # octal: drop a leading 0o and leading zeros
        v = v.removeprefix("0O").lstrip("0") or ("0" if v else "")
    elif qid == "q3":                    # binary with a fractional part
        v = v.replace(",", ".")
        if "." in v:
            i, f = v.split(".", 1)
            i = i.lstrip("0") or "0"
            f = f.rstrip("0")            # 1010.010 == 1010.01
            v = i + ("." + f if f else "")
        else:
            v = v.lstrip("0") or ("0" if v else "")
    elif qid == "q4":                    # two parts: binary | decimal
        parts = (v.split("|") + ["", ""])[:2]
        b = parts[0].lstrip("0") or ("0" if parts[0] else "")
        d = parts[1].lstrip("0") or ("0" if parts[1] else "")
        v = b + "|" + d
    return v


def decode(text):
    i = text.find(PREFIX)
    body = text[i + len(PREFIX):] if i >= 0 else text
    hx = re.sub(r"[^0-9A-Fa-f]", "", body)
    if len(hx) % 2 or len(hx) < 2 * (8 + 16 + 2):
        raise ValueError("code is incomplete or corrupted")
    raw = bytes.fromhex(hx)
    nonce, ct, tag = raw[:8], raw[8:-16], raw[-16:]
    pt = bytearray(len(ct))
    for blk in range((len(ct) + 31) // 32):
        ks = hashlib.sha256(SECRET + nonce + blk.to_bytes(4, "big")).digest()
        for j in range(32):
            k = blk * 32 + j
            if k >= len(ct):
                break
            pt[k] = ct[k] ^ ks[j]
    ok = hmac.compare_digest(hashlib.sha256(SECRET + nonce + bytes(pt)).digest()[:16], tag)
    if not ok:
        raise ValueError("integrity check failed (code was edited or copied incompletely)")
    return json.loads(pt.decode("utf-8")), nonce.hex()


def grade(data):
    ans, meta = data.get("answers", {}), data.get("meta", {})
    res, flags = {}, []
    total_pts = sum(POINTS.values())
    got = 0
    for q, correct in KEY.items():
        a = norm(q, ans.get(q, ""))
        if q == "q4":                     # two parts, half credit for one of them
            sa, sk = a.split("|"), norm(q, correct).split("|")
            frac = sum(1 for x, y in zip(sa, sk) if x == y and y) / 2
        else:
            frac = 1.0 if a == norm(q, correct) else 0.0
        ok = frac == 1.0
        res[q] = frac if frac in (0.0, 1.0) else round(frac, 2)
        got += POINTS[q] * frac
        if q in TRAPS and a == norm(q, TRAPS[q]):
            flags.append(f"TRAP_{q}")
        m = meta.get(q, {})
        if ok and m.get("t", 1e9) < FAST_SECONDS * 1000:
            flags.append(f"FAST_{q}({m['t']/1000:.1f}s)")
    blurs = sum(m.get("blurs", 0) for m in meta.values()) + data.get("global", {}).get("blurs", 0)
    blur_ms = sum(m.get("blurMs", 0) for m in meta.values()) + data.get("global", {}).get("blurMs", 0)
    pastes = sum(m.get("pastes", 0) for m in meta.values()) + data.get("global", {}).get("pastes", 0)
    copies = sum(m.get("copies", 0) for m in meta.values())
    refl = data.get("reflection", {})
    pastes += refl.get("meta", {}).get("pastes", 0)
    if blurs:
        det = [f"{q}:{m['blurs']}x,{m.get('blurMs',0)/1000:.0f}s"
               for q, m in meta.items() if m.get("blurs")]
        gb = data.get("global", {})
        if gb.get("blurs"):
            det.append(f"other:{gb['blurs']}x,{gb.get('blurMs',0)/1000:.0f}s")
        flags.append(f"FOCUS_LOSS({'; '.join(det)})")
    if pastes:
        flags.append(f"PASTE({pastes})")
    if copies:
        flags.append(f"QUESTION_COPY({copies})")
    if data.get("attempt", 1) != 1:
        flags.append(f"ATTEMPT={data.get('attempt')}")
    if data.get("ai") == "answer":
        flags.append("AI_DECLARED")
    return round(100 * got / total_pts, 1), res, flags, blurs, blur_ms, pastes



# ---------- minimal .xlsx writer/reader (standard library only) ----------
NUM_COLS = {*KEY, "score", "time_s", "focus_losses", "away_s", "pastes", "attempt"}
WIDTHS = {"file": 26, "valid": 7, "student_id": 13, "name": 22, "score": 8, "ai_review": 11, "flags": 60,
          "time_s": 8, "focus_losses": 13, "away_s": 8, "pastes": 8, "attempt": 8,
          "ai_declaration": 15, "started": 25, "submitted": 25, "explanation": 60, "code_id": 18}
_BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def question_marks(row):
    """Which question cells to highlight, from the flags text: 2 = strong, 1 = worth a look."""
    f = row.get("flags") or ""
    marks = {}
    for q in re.findall(r"TRAP_(q\d+)", f):
        marks[q] = 2
    for q in re.findall(r"FAST_(q\d+)", f):
        marks.setdefault(q, 1)
    m = re.search(r"FOCUS_LOSS\(([^)]*)\)", f)
    if m:
        for q, sec in re.findall(r"(q\d+):\d+x,(\d+)s", m.group(1)):
            if int(sec) >= 10:
                marks.setdefault(q, 1)
    return marks


def _col(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f.is_integer() else f


def write_xlsx(path, rows, cols):
    # styles: 0 normal, 1 header, 2 flagged cell, 3 score (0.0), 4 invalid row
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0"/></numFmts>'
        '<fonts count="3"><font><sz val="10"/><name val="Arial"/></font>'
        '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
        '<font><b/><sz val="10"/><name val="Arial"/></font></fonts>'
        '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1B2433"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD3E3F5"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF2A9A2"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="7">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>'
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>')
    out = []
    for i, c in enumerate(cols):
        out.append(f'<c r="{_col(i)}1" t="inlineStr" s="1"><is><t>{escape(c)}</t></is></c>')
    xml_rows = [f'<row r="1">{"".join(out)}</row>']
    for r_i, row in enumerate(rows, start=2):
        invalid = row.get("valid") == "NO"
        marks = question_marks(row)
        cells = []
        for c_i, c in enumerate(cols):
            v = row.get(c, "")
            if v is None or v == "":
                continue
            ref = f"{_col(c_i)}{r_i}"
            if invalid:
                st = 4
            elif c == "ai_review":
                st = 5 if v == "HIGH" else 6
            elif c in marks:
                st = 5 if marks[c] == 2 else 2
            elif c == "score":
                st = 3
            elif c == "flags":
                st = 2
            else:
                st = 0
            n = _num(v) if c in NUM_COLS else None
            if n is not None:
                cells.append(f'<c r="{ref}" s="{st}"><v>{n}</v></c>')
            else:
                txt = escape(_BAD_XML.sub("", str(v)))
                cells.append(f'<c r="{ref}" t="inlineStr" s="{st}"><is><t xml:space="preserve">{txt}</t></is></c>')
        xml_rows.append(f'<row r="{r_i}">{"".join(cells)}</row>')
    last = f"{_col(len(cols) - 1)}{max(1, len(rows) + 1)}"
    widths = "".join(f'<col min="{i+1}" max="{i+1}" width="{WIDTHS.get(c, 6)}" customWidth="1"/>'
                     for i, c in enumerate(cols))
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>'
        f'<cols>{widths}</cols><sheetData>{"".join(xml_rows)}</sheetData>'
        f'<autoFilter ref="A1:{last}"/></worksheet>')
    files = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        "xl/workbook.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Results" sheetId="1" r:id="rId1"/></sheets>'
            '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'
            f"Results!$A$1:${_col(len(cols)-1)}${max(1, len(rows)+1)}</definedName></definedNames></workbook>",
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>',
        "xl/worksheets/sheet1.xml": sheet,
        "xl/styles.xml": styles,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)


def read_xlsx(path):
    """Read the first sheet back into a list of dicts. Also handles files re-saved by Excel."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", ns):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{ns['m']}}}t")))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rid = wb.find("m:sheets/m:sheet", ns).get(f"{{{ns['r']}}}id")
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in rels if r.get("Id") == rid)
        target = target.lstrip("/")
        sheet_path = target if target.startswith("xl/") else "xl/" + target
        sheet = ET.fromstring(z.read(sheet_path))
    grid = []
    for row in sheet.iter(f"{{{ns['m']}}}row"):
        vals = {}
        for c in row.findall("m:c", ns):
            ref = re.match(r"([A-Z]+)", c.get("r")).group(1)
            idx = 0
            for ch in ref:
                idx = idx * 26 + ord(ch) - 64
            t = c.get("t")
            if t == "inlineStr":
                v = "".join(x.text or "" for x in c.iter(f"{{{ns['m']}}}t"))
            else:
                ve = c.find("m:v", ns)
                v = ve.text if ve is not None else ""
                if t == "s" and v != "":
                    v = shared[int(v)]
            vals[idx - 1] = v
        grid.append(vals)
    if not grid:
        return []
    header = [grid[0][i] for i in sorted(grid[0])]
    return [{h: g.get(i, "") for i, h in enumerate(header)} for g in grid[1:]]


def ai_review(row):
    """Turn the flags into one column: HIGH = strong AI signal, CHECK = worth a look."""
    f = row.get("flags") or ""
    if "TRAP_" in f or "SAME_CODE_IN_OTHER_FILE" in f or row.get("ai_declaration") == "answer":
        return "HIGH"
    away = float(row.get("away_s") or 0)
    hits = ("PASTE(" in f) + ("QUESTION_COPY(" in f) + ("SAME_EXPLANATION" in f) \
        + (f.count("FAST_") >= 2) + (away >= 20)
    return "CHECK" if hits else ""


CROSS_FLAGS = ("SAME_CODE_IN_OTHER_FILE", "SAME_EXPLANATION", "MULTIPLE_SUBMISSIONS")
COLS = ["file", "valid", "student_id", "name", *KEY, "score", "ai_review", "flags", "time_s", "focus_losses",
        "away_s", "pastes", "attempt", "ai_declaration", "started", "submitted", "explanation", "code_id"]


def row_key(r):
    return (r.get("code_id", ""), r.get("file", ""))


def grade_file(f, dump):
    row = {"file": str(f)}
    try:
        data, nonce = decode(f.read_text(encoding="utf-8", errors="ignore"))
    except Exception as e:
        row.update({"valid": "NO", "flags": f"UNDECODABLE: {e}"})
        return row
    score, res, flags, blurs, blur_ms, pastes = grade(data)
    st = data.get("student", {})
    t_total = sum(m.get("t", 0) for m in data.get("meta", {}).values()) / 1000
    row.update({
        "valid": "YES", "student_id": st.get("id", ""), "name": st.get("name", ""),
        **res, "score": score, "flags": " ".join(flags),
        "time_s": round(t_total), "focus_losses": blurs, "away_s": round(blur_ms / 1000),
        "pastes": pastes, "attempt": data.get("attempt"), "ai_declaration": data.get("ai"),
        "started": data.get("startedAt"), "submitted": data.get("submittedAt"),
        "explanation": data.get("reflection", {}).get("text", ""), "code_id": nonce,
    })
    if dump:
        Path(dump, f"{st.get('id','x')}_{nonce}.json").write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+")
    ap.add_argument("-o", "--out", default="quiz2B_results.xlsx")
    ap.add_argument("--dump", help="folder to write decoded JSON files")
    ap.add_argument("--fresh", action="store_true", help="ignore the existing results file and start over")
    args = ap.parse_args()

    files = []
    for p in map(Path, args.paths):
        files += sorted(p.rglob("*.txt")) if p.is_dir() else [p]
    files = [f.resolve() for f in files]
    if not files:
        sys.exit("No .txt files found.")
    if args.dump:
        Path(args.dump).mkdir(parents=True, exist_ok=True)

    # load previous results so new submissions are appended, not overwritten
    merged = {}
    out = Path(args.out)
    if out.exists() and not args.fresh:
        try:
            if out.suffix.lower() == ".csv":
                with open(out, newline="", encoding="utf-8-sig") as fh:
                    old = list(csv.DictReader(fh))
            else:
                old = read_xlsx(out)
        except PermissionError:
            sys.exit(f"Cannot read {out}. Close it in Excel and run again.")
        except (zipfile.BadZipFile, KeyError, StopIteration, ET.ParseError):
            sys.exit(f"{out} is not a results file this grader can read. Use a new -o name or --fresh.")
        for r in old:
            merged[row_key(r)] = r

    added = updated = 0
    this_run = []
    for f in files:
        r = grade_file(f, args.dump)
        k = row_key(r)
        if k in merged:
            updated += 1
        else:
            added += 1
        merged[k] = r
        this_run.append(k)

    # recompute cross-submission flags over the whole file
    rows = list(merged.values())
    valid = [r for r in rows if r.get("valid") == "YES"]
    norm_ex = lambda r: re.sub(r"\W+", "", (r.get("explanation") or "").lower())
    codes = Counter(r.get("code_id") for r in valid)
    exps = Counter(norm_ex(r) for r in valid)
    sids = Counter(r.get("student_id") for r in valid)
    for r in valid:
        base = [t for t in (r.get("flags") or "").split() if t not in CROSS_FLAGS]
        if codes[r.get("code_id")] > 1:
            base.append("SAME_CODE_IN_OTHER_FILE")
        if norm_ex(r) and exps[norm_ex(r)] > 1:
            base.append("SAME_EXPLANATION")
        if sids[r.get("student_id")] > 1:
            base.append("MULTIPLE_SUBMISSIONS")
        r["flags"] = " ".join(base)
    for r in rows:
        r["ai_review"] = ai_review(r) if r.get("valid") == "YES" else ""

    try:
        if out.suffix.lower() == ".csv":
            with open(out, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.DictWriter(fh, fieldnames=COLS, extrasaction="ignore")
                w.writeheader()
                w.writerows(rows)
        else:
            write_xlsx(out, rows, COLS)
    except PermissionError:
        sys.exit(f"Cannot write {out}. Close it in Excel and run again.")

    print(f"This run: {len(files)} file(s), {added} new, {updated} re-graded")
    print(f"{out} now has {len(rows)} row(s), {len(valid)} valid")
    if valid:
        print(f"Class average: {sum(float(r['score']) for r in valid)/len(valid):.1f}")
    for k in this_run:
        r = merged[k]
        if r.get("flags"):
            print(f"  {r.get('student_id') or '?':>12}  {str(r.get('score','-')):>5}  {r['flags']}")


if __name__ == "__main__":
    main()
