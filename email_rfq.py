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


def build_filled_table(lines, filled_path, output_path):
    """Rebuild the output as: original columns + Price + Description + IMPA."""
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

    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    headers = ["Line No", "Stk ID", "Name", "Qty", "UOM", "Price", "Description", "IMPA"]
    for c, h in enumerate(headers, start=1):
        ws2.cell(row=1, column=c).value = h
    for i, ln in enumerate(lines):
        r = 2 + i
        f = filled.get(str(i + 1), {})
        ws2.cell(row=r, column=1).value = ln.get("line_no")
        ws2.cell(row=r, column=2).value = ln.get("stk") or None
        ws2.cell(row=r, column=3).value = ln.get("name")
        ws2.cell(row=r, column=4).value = ln.get("qty")
        ws2.cell(row=r, column=5).value = ln.get("uom")
        ws2.cell(row=r, column=6).value = f.get("price")
        ws2.cell(row=r, column=7).value = f.get("remark")
        # IMPA = the customer's Stk ID if present, otherwise the app's parsed/matched guess
        impa = (ln.get("stk") or f.get("impa"))
        ws2.cell(row=r, column=8).value = impa if impa not in (None, "") else None
    wb2.save(output_path)
    return output_path


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
    # rebuild output as the original email columns + Price + Description + IMPA
    build_filled_table(lines, output_path, output_path)
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
