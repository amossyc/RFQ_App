"""Parse an email-list RFQ (a table in the email body) into the portal-export Excel
layout, then run the existing fill_quote logic on it — fully offline (no browser,
no portal, no Outlook)."""

import re, html, os, sys, glob
import openpyxl
import fill_quote
import config


def clean_cell(td):
    t = re.sub(r"<[^>]+>", "", td)
    t = html.unescape(t)
    return re.sub(r"\s+", " ", t).strip()


def _field_for(header_text):
    """Map a header cell text to a field name (handles all table layouts)."""
    h = (header_text or "").lower().strip("*").strip()
    if "line no" in h or h in ("line", "no", "no.", "s/no", "s.no", "sn", "item", "item no"):
        return "line_no"
    if "line type" in h:
        return "line_type"
    if "stk" in h or "stock" in h:
        return "stk"
    if "name" in h:
        return "name"
    if "qty" in h or "quantity" in h:
        return "qty"
    if "uom" in h or "unit" in h:
        return "uom"
    return None


def _infer_headerless_layout(parsed):
    """Infer column roles for a headerless table from the first few data rows."""
    sample = [r for r in parsed if len(r) >= 4][:6]
    if not sample:
        return None
    ncols = max(len(r) for r in sample)
    if ncols < 4:
        return None

    cols = []
    for j in range(ncols):
        vals = [str(r[j]).strip() for r in sample if j < len(r) and str(r[j]).strip()]
        cols.append(vals)

    def is_num(s):
        try:
            float(s.replace(",", ""))
            return True
        except Exception:
            return False

    def is_uom(s):
        return bool(re.fullmatch(r"[A-Za-z.]{1,5}", s))

    def is_code(s):
        return bool(re.fullmatch(r"\d{1,3}-\d{1,3}-\d{1,3}", s)) or bool(re.fullmatch(r"\d{5,7}", s))

    layout = {"line_no": 0}
    for j in range(1, ncols):
        vals = cols[j]
        if not vals:
            continue
        if all(is_uom(v) for v in vals):
            layout.setdefault("uom", j)
        elif all(is_code(v) for v in vals):
            layout.setdefault("stk", j)
        elif all(is_num(v) for v in vals):
            layout.setdefault("qty", j)

    used = set(layout.values())
    for j in range(1, ncols):
        if j not in used:
            layout.setdefault("name", j)
            break

    if {"line_no", "name", "qty", "uom"} <= set(layout):
        return layout
    return None


def parse_email_table(html_body):
    """Parse an RFQ table from the email HTML body into a list of line dicts.

    Column order is detected from the header row, so both the 6-column
    (Line No / Line Type / Stk ID / Name / Qty / UOM) and the 5-column
    (Line No / Stk ID / Qty / UOM / Name) layouts work."""
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", html_body or "", re.IGNORECASE | re.DOTALL)
    parsed = []
    for r in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", r, re.IGNORECASE | re.DOTALL)
        parsed.append([clean_cell(c) for c in cells])

    # find the header row and map columns
    header_map = {}
    header_idx = None
    for i, vals in enumerate(parsed):
        fields = [_field_for(v) for v in vals]
        if "line_no" in fields and "name" in fields:
            header_idx = i
            for j, f in enumerate(fields):
                if f:
                    header_map[f] = j
            break
    if not header_map:
        # headerless table: infer the columns from the data
        header_map = _infer_headerless_layout(parsed) or {}
        if header_map:
            header_idx = -1
        else:
            return []

    def get(vals, field):
        j = header_map.get(field)
        return vals[j] if j is not None and j < len(vals) else ""

    lines = []
    for vals in parsed[header_idx + 1:]:
        line_no = get(vals, "line_no")
        name = get(vals, "name")
        if not name or not line_no:
            continue
        lines.append({
            "line_no": line_no,
            "line_type": get(vals, "line_type"),
            "stk": get(vals, "stk"),
            "name": name,
            "qty": get(vals, "qty"),
            "uom": get(vals, "uom"),
        })
    return lines


def _get_table_header(html_body):
    """Return (header_cells, field_order) from the email table's header row."""
    for r in re.findall(r"<tr[^>]*>(.*?)</tr>", html_body or "", re.IGNORECASE | re.DOTALL):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", r, re.IGNORECASE | re.DOTALL)
        vals = [clean_cell(c) for c in cells]
        fields = [_field_for(v) for v in vals]
        if "line_no" in fields and "name" in fields:
            return vals, fields
    return [], []


def extract_quality_remark(html_body):
    """Pull a quality instruction like '== QUOTE IRENO OR SAME AS REMARK ==' from the body."""
    text = re.sub(r"<[^>]+>", " ", html_body or "")
    text = html.unescape(text)
    m = re.search(r"==\s*QUOTE\s+(.*?)\s*==", text, re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else ""


def build_portal_xlsx(lines, path, quality=""):
    """Build an Excel in the portal-export layout (same columns fill_quote expects)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    if quality:
        ws.cell(row=1, column=5).value = "Remarks:"
        ws.cell(row=1, column=6).value = quality
    for i, ln in enumerate(lines):
        r = 2 + i
        ws.cell(row=r, column=1).value = r                     # A: counter
        ws.cell(row=r, column=3).value = str(i + 1)            # C: line id
        ws.cell(row=r, column=4).value = ln["line_no"]         # D: line number
        ws.cell(row=r, column=5).value = ln["stk"] or None     # E: Stk ID
        ws.cell(row=r, column=6).value = ln["name"]            # F: Name
        ws.cell(row=r, column=7).value = ln["qty"]             # G: Qty
        ws.cell(row=r, column=8).value = ln["uom"]             # H: UOM
    wb.save(path)
    return path


def build_filled_table(html_body, lines, filled_path, output_path):
    """Rebuild the output as: the customer's original table columns + Price + Description + IMPA."""
    from openpyxl.styles import Border, Side, Font
    # read filled values (price/remark/impa) from the portal-format filled file
    wb = openpyxl.load_workbook(filled_path, data_only=True)
    ws = wb.active
    filled = {}
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=3).value in (None, ""):
            continue
        key = str(ws.cell(row=r, column=3).value)
        filled[key] = {
            "price": ws.cell(row=r, column=9).value,
            "remark": ws.cell(row=r, column=11).value,
            "impa": ws.cell(row=r, column=12).value,
        }
    wb.close()

    header_cells, field_order = _get_table_header(html_body)

    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    headers = list(header_cells) + ["Price", "Description", "IMPA"]
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(headers, start=1):
        cell = ws2.cell(row=1, column=c, value=h)
        cell.border = border
        cell.font = Font(bold=True)
    for i, ln in enumerate(lines):
        r = 2 + i
        f = filled.get(str(i + 1), {})
        vals = [ln.get(field) if field else "" for field in field_order]
        vals.append(f.get("price"))
        vals.append(f.get("remark"))
        impa = (ln.get("stk") or f.get("impa"))
        vals.append(impa if impa not in (None, "") else None)
        for c, v in enumerate(vals, start=1):
            cell = ws2.cell(row=r, column=c, value=v)
            cell.border = border
    wb2.save(output_path)
    return output_path


def _read_filled_values(filled_path):
    """Read {price, remark, impa} per line from the portal-format filled file."""
    wb = openpyxl.load_workbook(filled_path, data_only=True)
    ws = wb.active
    out = []
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=3).value in (None, ""):
            continue
        stk = ws.cell(row=r, column=5).value
        parsed = ws.cell(row=r, column=12).value
        impa = stk or parsed
        out.append({
            "price": ws.cell(row=r, column=9).value,
            "remark": ws.cell(row=r, column=11).value,
            "impa": impa if impa not in (None, "") else None,
        })
    wb.close()
    return out


def build_html_table(html_body, filled_values):
    """Add 3 columns (Price, Description, IMPA) to the customer's ORIGINAL HTML table."""
    m = re.search(r"<table[^>]*>.*?</table>", html_body, re.IGNORECASE | re.DOTALL)
    if not m:
        return None
    table = m.group(0)

    style = "style='border:solid gray 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'"

    def cell_html(value, bold=False):
        v = html.escape(str(value)) if value is not None else "&nbsp;"
        inner = ("<b>%s</b>" % v) if bold else v
        return "<td nowrap valign=bottom %s>%s</td>" % (style, inner)

    state = {"header": False, "data_i": 0}

    def transform(match):
        tr = match.group(0)
        cells = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.IGNORECASE | re.DOTALL)
        fields = [_field_for(clean_cell(c)) for c in cells]
        if not state["header"] and "line_no" in fields and "name" in fields:
            state["header"] = True
            add = cell_html("Price", True) + cell_html("Description", True) + cell_html("IMPA", True)
        elif state["header"]:
            i = state["data_i"]
            f = filled_values[i] if i < len(filled_values) else {}
            state["data_i"] += 1
            add = cell_html(f.get("price")) + cell_html(f.get("remark")) + cell_html(f.get("impa"))
        else:
            return tr
        idx = tr.lower().rfind("</tr>")
        return tr[:idx] + add + tr[idx:]

    return re.sub(r"<tr[^>]*>.*?</tr>", transform, table, flags=re.IGNORECASE | re.DOTALL)


def build_clean_html_table(lines, filled_values):
    """Build a simple, clean HTML table (separate from the customer's table)."""
    hdr = ["Line No", "Stk ID", "Name", "Qty", "UOM", "Price", "Description", "IMPA"]
    out = ['<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:11px">']
    out.append("<tr>" + "".join("<th style='background:#eeeeee'>%s</th>" % h for h in hdr) + "</tr>")
    for i, ln in enumerate(lines):
        f = filled_values[i] if i < len(filled_values) else {}
        impa = ln.get("stk") or f.get("impa") or ""
        price = f.get("price") if f.get("price") is not None else ""
        remark = f.get("remark") or ""
        cells = [ln.get("line_no") or "", ln.get("stk") or "", ln.get("name") or "",
                 ln.get("qty") or "", ln.get("uom") or "", price, remark, impa]
        out.append("<tr>" + "".join("<td>%s</td>" % html.escape(str(c)) for c in cells) + "</tr>")
    out.append("</table>")
    return "".join(out)


def newest_pricelist():
    files = [f for f in glob.glob(os.path.join(config.PRICELIST_DIR, "*.xlsm"))
             if not os.path.basename(f).startswith("~$")]
    return sorted(files, key=os.path.getmtime)[-1] if files else None


def process_email_rfq(html_body, output_path, pricelist_path=None):
    lines = parse_email_table(html_body)
    if not lines:
        print("No table lines found in the email.")
        return None
    print(f"Parsed {len(lines)} lines.")
    pl = pricelist_path or newest_pricelist()
    if not pl:
        raise SystemExit("No pricelist found.")
    print("Pricelist:", pl)
    tmp = output_path + ".tmp.xlsx"
    quality = extract_quality_remark(html_body)
    build_portal_xlsx(lines, tmp, quality=quality)
    rep = fill_quote.fill_quote(tmp, pl, output_path)
    try:
        os.remove(tmp)
    except Exception:
        pass
    # read the filled values, then rebuild the output as the customer's table + 3 columns
    filled_values = _read_filled_values(output_path)
    build_filled_table(html_body, lines, output_path, output_path)
    rep["html_table"] = build_clean_html_table(lines, filled_values)
    rep["lines"] = len(lines)
    return rep


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python email_rfq.py <html_or_dump_file> [output.xlsx]")
        raise SystemExit(1)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + "-done.xlsx"
    with open(src, encoding="utf-8") as f:
        content = f.read()
    if "HTML:" in content:
        content = content.split("HTML:", 1)[1]  # it's a dump file
    rep = process_email_rfq(content, out)
    if rep:
        print("Done ->", out)
        print(f"total={rep.get('total')} flagged={len(rep.get('flagged', []))} "
              f"auto_matched={len(rep.get('auto_matched', []))}")
