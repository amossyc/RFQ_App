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
    """Map a header cell text to a field name (handles both table layouts)."""
    h = (header_text or "").lower().strip("*").strip()
    if "line no" in h or h == "line":
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


def build_portal_xlsx(lines, path):
    """Build an Excel in the portal-export layout (same columns fill_quote expects)."""
    wb = openpyxl.Workbook()
    ws = wb.active
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
    build_portal_xlsx(lines, tmp)
    rep = fill_quote.fill_quote(tmp, pl, output_path)
    try:
        os.remove(tmp)
    except Exception:
        pass
    # read the filled values, then rebuild the output as the customer's table + 3 columns
    filled_values = _read_filled_values(output_path)
    build_filled_table(html_body, lines, output_path, output_path)
    rep["html_table"] = build_html_table(html_body, filled_values)
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
