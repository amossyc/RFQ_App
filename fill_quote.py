import openpyxl, re, os, glob, difflib
import config

WORK = config.WORK
PRICELIST_DIR = config.PRICELIST_DIR

GOOD = {"GERMANY", "UK/GERMANY", "BMI GERMANY", "BOSCH", "BAHCO", "PYRAMEX", "PRO-VISE", "IMA ITALY"}
AVG = {"FAPRO", "STANLEY"}
SPECIAL_KEYS = ("STAINLESS", "S.STEEL", "NON-SPARK", "NONSPARK", "NO-SPARK", "NOSPARK", "NON SPARK")

# 6 digits, with optional spaces or dashes between digit groups: 616623 / 61 66 23 / 61-66-23
CODE_RE = re.compile(r"\b\d(?:[\s\-]*\d){5}\b")
# "IMPA" directly attached to the code (no space/hyphen): IMPA612832 / IMPA 612832 / IMPA-612832
IMPA_RE = re.compile(r"IMPA[\s\-\:\.]*(\d(?:[\s\-]*\d){5})", re.IGNORECASE)
# "IMPA CODE:B850601" / "IMPA CODE: 850601" (with optional letter prefix)
IMPA_CODE_RE = re.compile(r"IMPA\s*(?:CODE\s*)?:?\s*[A-Za-z]?(\d(?:[\s\-]*\d){5})", re.IGNORECASE)

def num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def load_pricelist(path):
    data = {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["SINWA"]
    nocode = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        brands = []
        if row[9] not in (None, ""):
            brands.append((str(row[9]).strip(), None))  # primary brand (price in tender/net cols)
        for b in range(11, len(row), 2):  # L, N, P, ... -> alternatives, price in preceding col
            brand = row[b]
            if brand not in (None, ""):
                brands.append((str(brand).strip(), row[b - 1]))
        entry = {"tender": row[6], "tender25": row[7], "net": row[8], "brands": brands, "desc": row[3]}
        impa = row[0]
        if impa in (None, ""):
            nocode += 1
            entry["no_code"] = True
            data[f"NOCODE_{nocode}"] = entry
        else:
            data[str(impa).strip()] = entry
    wb.close()
    return data

def pick_price(e):
    for k in ("tender25", "tender", "net"):
        n = num(e.get(k))
        if n is not None:
            return n
    return 0.0

def tier(brand):
    b = brand.upper()
    if b in GOOD:
        return "good"
    if b in AVG:
        return "avg"
    return "cheap"

def detect_quality(text):
    t = (text or "").upper()
    if "GOOD" in t:
        return "good"
    if "AVERAGE" in t or "FAPRO" in t or "IRENO" in t:
        return "avg"
    return None

def is_special(text):
    t = (text or "").upper()
    if any(k in t for k in SPECIAL_KEYS):
        return True
    # "SS" as a standalone token = stainless steel (avoid HSS, BRASS, CLASS, etc.)
    for tok in re.split(r"[^A-Z0-9]+", t):
        if tok == "SS":
            return True
    return False

def fmt_price(v):
    f = float(v)
    if f.is_integer():
        return f"${int(f)}"
    return f"${f:.2f}"

def extract_impa_codes(name):
    """Return 6-digit IMPA candidates found anywhere in the description."""
    if not name:
        return []
    out = []
    for m in CODE_RE.finditer(name):
        code = re.sub(r"[\s\-]", "", m.group(0))
        if len(code) == 6 and code != "999999" and code not in out:
            out.append(code)
    # also match "IMPA612832" / "IMPA 612832" / "IMPA-612832" (code attached to "IMPA")
    for m in IMPA_RE.finditer(name):
        code = re.sub(r"[\s\-]", "", m.group(1))
        if len(code) == 6 and code != "999999" and code not in out:
            out.append(code)
    # "IMPA CODE:B850601" / "IMPA CODE: 850601" (label + optional letter prefix)
    for m in IMPA_CODE_RE.finditer(name):
        code = re.sub(r"[\s\-]", "", m.group(1))
        if len(code) == 6 and code != "999999" and code not in out:
            out.append(code)
    return out

_PUNCT = ".,;:()[]{}'\"!?<>|/\\"

def _norm(w):
    return w.replace("-", "")

def _compact(s):
    return (s or "").replace("-", "").replace(" ", "").upper()

def _words(s):
    out = []
    for x in re.split(r"[,\s;|]+", (s or "").upper()):
        x = x.strip(_PUNCT)
        if x:
            out.append(x)
    return out

def brand_in_description(name, brands):
    """Return (brand_cell, display_name) if the description names a brand, else None."""
    up = (name or "").upper()
    up_words = _words(up)
    up_norm = {_norm(w): w for w in up_words}
    for b, p in brands:
        if not b:
            continue
        bup = b.upper()
        if _norm(bup) in _norm(up):
            return (b, b)
        for w in _words(bup):
            if len(w) < 4 or any(c.isdigit() for c in w):
                continue
            if w in GENERIC_WORDS:
                continue
            wn = _norm(w)
            if wn in up_norm:
                return (b, w)
            for uw in up_words:
                if len(uw) >= 4 and abs(len(wn) - len(_norm(uw))) <= 1:
                    if difflib.SequenceMatcher(None, wn, _norm(uw)).ratio() > 0.8:
                        return (b, w)
    return None

GENERIC_WORDS = {
    "ELECTRIC", "ELECTRICAL", "MIN", "MAX", "HIGH", "LOW", "PRESSURE",
    "WITH", "FOR", "THE", "AND", "OF", "IN", "ON", "OR", "A", "AN",
    "SIZE", "TYPE", "MODEL", "GRADE", "STANDARD", "HEAVY", "LIGHT", "DUTY",
    "PNEUMATIC", "HYDRAULIC", "AUTOMATIC", "MANUAL", "POWER", "INDUSTRIAL",
    "DIA", "HEAD", "BODY",
    "SPLIETHOFF", "GROUP", "CATALOG", "REF",
}

SYNONYMS = {
    "PHILLIPS": "PHILLIP",
    "PHILLIP": "PHILLIP",
}

def _meaningful_tokens(text):
    core = re.split(r"\s*\|\s*(?:IMPA|SSM|Ref\.)", text or "", maxsplit=1)[0]
    # normalize "NO.2" / "NO 2" -> "NO2" (Phillips / size number)
    core = re.sub(r"NO\.\s*(\d)", r"NO\1", core, flags=re.IGNORECASE)
    core = re.sub(r"[^A-Za-z0-9. ]", " ", core).upper()
    out = []
    for t in core.split():
        t = t.strip(".")
        if len(t) < 3:
            continue
        if any(c.isdigit() for c in t) and "MM" not in t and not re.match(r"NO\d+$", t):
            continue
        if t in GENERIC_WORDS:
            continue
        out.append(SYNONYMS.get(t, t))
    return out

def _size_tokens(text):
    t = (text or "").upper()
    out = set()
    # ranges: 10-32MM / 10-32 MM / 10 - 32 MM
    for m in re.findall(r"\d+(?:\.\d+)?\s*-\s*\d+(?:\.\d+)?\s*MM", t):
        a, b = re.findall(r"\d+(?:\.\d+)?", m)
        out.add(f"{min(float(a), float(b)):g}-{max(float(a), float(b)):g}MM")
    # single sizes and pairs: 14MM / 14 MM / 14X17MM / 14 X 17 MM
    for m in re.findall(r"\d+(?:\.\d+)?(?:\s*[Xx]\s*\d+(?:\.\d+)?)?\s*MM", t):
        body = m[:-2].strip()
        if "X" in body:
            a, b = body.upper().split("X")
            out.add(f"{min(float(a), float(b))}X{max(float(a), float(b))}MM")
        else:
            out.add(f"{float(body):g}MM")
    return out

PRIORITY_FEATURES = ("NONSPARK", "STAINLESS", "100V")

def _priority_features(text):
    c = _compact(text)
    return [f for f in PRIORITY_FEATURES if f in c]

def _voltage(text):
    """Voltage specs like 1000V, 220V, 110V — a hard requirement when the customer specifies one."""
    return set(re.findall(r"\b\d+V\b", (text or "").upper()))

def _rank_variant(desc):
    """Tie-break rank for a pricelist description: DEEP > 12-PT > 6-PT > other."""
    d = (desc or "").upper()
    if "DEEP" in d:
        return 3
    if "12-PT" in d or "12PT" in d or "12 POINT" in d or "12POINT" in d:
        return 2
    if "6-PT" in d or "6PT" in d or "6 POINT" in d or "6POINT" in d:
        return 1
    return 0

def _leading_match(name, desc):
    """Count leading meaningful tokens of name that match desc (prefix-style)."""
    nt = _meaningful_tokens(name)
    dt = _meaningful_tokens(desc)
    n = 0
    for a, b in zip(nt, dt):
        if a == b:
            n += 1
        else:
            break
    return n

def match_by_description(name, data):
    """Conservative fallback: match a no-code item to a pricelist row by name.
    Priority features (NON-SPARK / STAINLESS / 100V) must match; size gives a strong
    bonus; otherwise shared meaningful tokens decide."""
    # AUTO DARK feature -> match the item whose brands carry 'AUTO DARK'
    if "AUTODARK" in _compact(name):
        m = [impa for impa, e in data.items() if "AUTODARK" in _compact(" ".join(b for b, _ in e.get("brands", [])))]
        if len(m) == 1:
            return m[0]
    feats = _priority_features(name)
    volts = _voltage(name)
    tokens = _meaningful_tokens(name)
    sizes = _size_tokens(name)
    scored = []
    for impa, e in data.items():
        desc = str(e.get("desc") or "")
        if feats and not all(f in _priority_features(desc) for f in feats):
            continue
        if volts and not (volts & _voltage(desc)):
            continue  # customer specified a voltage; item must match it
        dtokens = set(_meaningful_tokens(desc))
        overlap = sum(1 for t in tokens if t in dtokens)
        size_overlap = len(sizes & _size_tokens(desc)) if sizes else 0
        if overlap + (1 if size_overlap else 0) < 3:
            continue  # need enough evidence: shared type words + a size match
        size_bonus = 10 * size_overlap if overlap >= 3 else 0
        scored.append((overlap + size_bonus, impa))
    if not scored:
        return None
    scored.sort(key=lambda x: -x[0])
    if len(scored) == 1 or scored[0][0] > scored[1][0]:
        return scored[0][1]
    # tie-break among top-scored items: prefer 12-PT > 6-PT; DEEP only when requested;
    # then prefer the item whose leading words match the customer's wording
    top_score = scored[0][0]
    tied = [impa for s, impa in scored if s == top_score]
    wants_deep = "DEEP" in (name or "").upper()
    ranked = []
    for impa in tied:
        desc = str(data[impa].get("desc") or "")
        r = _rank_variant(desc)
        if r == 3 and not wants_deep:
            continue
        ranked.append((r, _leading_match(name, desc), impa))
    if not ranked:
        return None
    ranked.sort(key=lambda x: (-x[0], -x[1]))
    if len(ranked) == 1 or (ranked[0][0], ranked[0][1]) != (ranked[1][0], ranked[1][1]):
        return ranked[0][2]
    return None

def mentioned_brand(name):
    """Extract the brand the customer named (the word after 'RefernenceNo : CODE,')."""
    m = re.search(r"RefernenceNo\s*:\s*\d+\s*[,\s]\s*([A-Za-z0-9\-\./]+)", name or "", re.IGNORECASE)
    return m.group(1).upper() if m else None

def quote_item(entry, name, quality):
    """Return (price, remark) for a line item, honoring any brand the customer named."""
    brands = entry["brands"]
    primary_price = pick_price(entry)
    if not brands:
        return primary_price, ""
    if "COBALT" not in (name or "").upper():
        brands = [(b, p) for b, p in brands if "COBALT" not in b.upper()]
    if not brands:
        return primary_price, ""
    primary = brands[0][0]

    # 0. AUTO DARK feature -> select auto-dark variants (customer wants auto-dark helmet etc.)
    if "AUTODARK" in _compact(name):
        auto = [(b, num(p)) for b, p in brands if "AUTODARK" in _compact(b) and num(p) is not None]
        if auto:
            auto.sort(key=lambda x: x[1])
            budget = auto[0]
            premium = auto[-1] if len(auto) >= 2 else None
            if quality == "good" and premium is not None:
                return budget[1], f"{budget[0]} / {fmt_price(premium[1])} {premium[0]}"
            return budget[1], budget[0]

    # 1. stainless / non-spark -> two options (by description OR by the item's primary brand)
    if (is_special(name) or is_special(primary)) and len(brands) >= 2:
        alt = brands[1]
        if num(alt[1]) is not None:
            return primary_price, f"{primary} / {fmt_price(alt[1])} {alt[0]}"

    # 2. explicit brand named in the description
    found = brand_in_description(name, brands)
    if found:
        brand_cell, display = found
        named_price = None
        for b, p in brands:
            if b == brand_cell:
                named_price = num(p)
                break
        if named_price is not None:
            # named an ALTERNATIVE -> quote at its price, primary is the cheap option
            return named_price, f"{display} / {fmt_price(primary_price)} {primary}"
        else:
            # named the PRIMARY -> quote at tender, first alternative is the option
            alt = brands[1] if len(brands) >= 2 else None
            if alt and num(alt[1]) is not None:
                return primary_price, f"{brand_cell} / {fmt_price(alt[1])} {alt[0]}"
            return primary_price, brand_cell

    # 3. brand mentioned but not in our pricelist -> offer our standard two options
    if mentioned_brand(name) and len(brands) >= 2 and num(brands[1][1]) is not None:
        return primary_price, f"{primary} / {fmt_price(brands[1][1])} {brands[1][0]}"

    # 4. quality override -> quote the selected tier's brand AT ITS OWN PRICE
    if quality == "good":
        for b, p in brands:
            if tier(b) == "good":
                return (num(p) if num(p) is not None else primary_price), b
        return primary_price, primary
    if quality == "avg":
        for b, p in brands:
            if tier(b) == "avg":
                return (num(p) if num(p) is not None else primary_price), b
        # no FAPRO/avg brand -> cheap + next-cheapest alternative
        alts = [(b, num(p)) for b, p in brands[1:] if num(p) is not None]
        if alts:
            cheap = min(alts, key=lambda x: x[1])
            return primary_price, f"{primary} / {fmt_price(cheap[1])} {cheap[0]}"
        return primary_price, primary

    # 5. default: primary only
    return primary_price, primary

def fill_quote(input_path, pricelist_path, output_path):
    data = load_pricelist(pricelist_path)
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    top_remarks = ""
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=5).value == "Remarks:":
            v = ws.cell(row=r, column=6).value or ""
            if v:
                top_remarks = v  # keep the LAST (actual) Remarks field, not the template placeholder

    total = 0
    flagged = []
    auto = []
    for r in range(1, ws.max_row + 1):
        line_id = ws.cell(row=r, column=3).value
        if line_id in (None, ""):
            continue
        line_no = ws.cell(row=r, column=4).value
        line_ref = str(line_no).strip() if line_no not in (None, "") else str(line_id)
        total += 1
        name = ws.cell(row=r, column=6).value or ""
        if "CHK UOM" in name.upper():
            flagged.append((line_ref, str(ws.cell(row=r, column=5).value or "").strip(), "[CHK UOM] " + name[:45]))
            ws.cell(row=r, column=9).value = None
            ws.cell(row=r, column=11).value = None
            continue
        stk = ws.cell(row=r, column=5).value
        impa = str(stk).strip() if stk not in (None, "") else ""
        if impa == "":
            codes = extract_impa_codes(name)
            for code in codes:
                if code in data:
                    ws.cell(row=r, column=5).value = code
                    impa = code
                    break
            if impa == "" and codes:
                ws.cell(row=r, column=5).value = codes[0]
                impa = codes[0]
        entry = data.get(impa)
        if entry is not None:
            # sanity check: does the customer's description match this IMPA's pricelist row?
            desc = str(entry.get("desc") or "")
            if desc.strip():
                impa_tokens = set(_meaningful_tokens(desc))
                name_tokens = _meaningful_tokens(name)
                overlap = sum(1 for t in name_tokens if t in impa_tokens)
                size_overlap = len(_size_tokens(name) & _size_tokens(desc))
                volts = _voltage(name)
                voltage_mismatch = bool(volts) and not (volts & _voltage(desc))
                match_ratio = overlap / len(name_tokens) if name_tokens else 1.0
                weak_match = len(name_tokens) >= 4 and match_ratio < 0.4
                if (overlap == 0 and size_overlap == 0) or voltage_mismatch or weak_match:
                    flagged.append((line_ref, impa, "[IMPA mismatch] " + name[:45]))
                    ws.cell(row=r, column=9).value = 0
                    ws.cell(row=r, column=11).value = "IMPA MISMATCH @@@@@@"
                    continue
        if entry is None:
            matched = match_by_description(name, data)
            if matched:
                if str(matched).startswith("NOCODE_"):
                    impa = ""
                    entry = data.get(matched)
                    auto.append((line_ref, "(no IMPA code)", name[:45]))
                else:
                    ws.cell(row=r, column=5).value = matched
                    impa = matched
                    entry = data.get(matched)
                    auto.append((line_ref, matched, name[:45]))
        if entry is None:
            flagged.append((line_ref, impa, name[:45]))
            ws.cell(row=r, column=9).value = None
            ws.cell(row=r, column=11).value = None
            continue
        quality = detect_quality(top_remarks) or detect_quality(name)
        price, remark = quote_item(entry, name, quality)
        ws.cell(row=r, column=9).value = price
        if remark:
            ws.cell(row=r, column=11).value = remark
    wb.save(output_path)
    return {"total": total, "flagged": flagged, "pricelist": len(data), "auto_matched": auto}

def newest(pattern):
    files = glob.glob(pattern)
    if not files:
        return None
    return max(files, key=os.path.getmtime)

if __name__ == "__main__":
    rfq = newest(os.path.join(WORK, "RFQ-*.xlsx"))
    pl = newest(os.path.join(PRICELIST_DIR, "*.xlsm"))
    if not rfq:
        print("No RFQ file found in", WORK)
    elif not pl:
        print("No pricelist found in", PRICELIST_DIR)
    else:
        out = os.path.join(WORK, os.path.basename(rfq).replace(".xlsx", "-done.xlsx"))
        print("RFQ:", rfq)
        print("Pricelist:", pl, "(", os.path.getsize(pl), "bytes )")
        rep = fill_quote(rfq, pl, out)
        print("Saved:", out)
        print("Pricelist rows:", rep["pricelist"])
        print("Line items:", rep["total"])
        print("Flagged (no pricelist match):", len(rep["flagged"]))
        for f in rep["flagged"]:
            print("   FLAG", f)
