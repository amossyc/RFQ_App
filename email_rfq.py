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


def parse_email_table(html_body):
    """Parse an RFQ table from the email HTML body into a list of line dicts."""
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", html_body or "", re.IGNORECASE | re.DOTALL)
    lines = []
    for r in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", r, re.IGNORECASE | re.DOTALL)
        if len(cells) < 6:
            continue
        vals = [clean_cell(c) for c in cells]
        line_no, line_type, stk, name, qty, uom = vals[:6]
        if not name:
            continue
        if line_no.lower().startswith("line"):
            continue  # header row
        lines.append({
            "line_no": line_no,
            "line_type": line_type,
            "stk": stk,
            "name": name,
            "qty": qty,
            "uom": uom,
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
