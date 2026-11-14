import os
import sys
import sqlite3
import datetime
from collections import defaultdict
from statistics import median
import openpyxl
import config

DB = config.DB
REPORTS_DIR = config.REPORTS_DIR

def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def init_db(db_path=DB):
    conn = sqlite3.connect(db_path)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS quotes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rfq_ref TEXT,
            subject TEXT,
            line_id TEXT,
            impa TEXT,
            description TEXT,
            qty REAL,
            uom TEXT,
            price REAL,
            brand TEXT,
            quoted_at TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_impa ON quotes(impa)")
    conn.commit()
    return conn

def read_filled_lines(filled_path):
    """Extract the quoted line items from a FILLED workbook (skips blank/flagged lines)."""
    wb = openpyxl.load_workbook(filled_path, data_only=True)
    ws = wb.active
    lines = []
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=3).value in (None, ""):
            continue
        price = _num(ws.cell(row=r, column=9).value)
        if price is None:
            continue
        lines.append({
            "line_id": str(ws.cell(row=r, column=3).value),
            "impa": str(ws.cell(row=r, column=5).value or "").strip(),
            "description": str(ws.cell(row=r, column=6).value or "").strip(),
            "qty": _num(ws.cell(row=r, column=7).value),
            "uom": str(ws.cell(row=r, column=8).value or "").strip(),
            "price": price,
            "brand": str(ws.cell(row=r, column=11).value or "").strip(),
        })
    wb.close()
    return lines

def log_quote(db_path, rfq_ref, subject, lines, quoted_at=None):
    conn = init_db(db_path)
    ts = quoted_at or datetime.datetime.now().isoformat(timespec="seconds")
    for ln in lines:
        conn.execute(
            "INSERT INTO quotes (rfq_ref, subject, line_id, impa, description, qty, uom, price, brand, quoted_at) "
            "VALUES (?,?,?,?,?,?,?,?,?,?)",
            (rfq_ref, subject, ln["line_id"], ln["impa"], ln["description"], ln["qty"], ln["uom"], ln["price"], ln["brand"], ts),
        )
    conn.commit()
    conn.close()

def generate_report(db_path=DB, output_path=None, period_days=30):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM quotes ORDER BY quoted_at").fetchall()
    conn.close()
    if not rows:
        print("No quote data yet — nothing to report.")
        return None

    now = datetime.datetime.now()
    cutoff = (now - datetime.timedelta(days=period_days)).isoformat(timespec="seconds")

    prices = defaultdict(list)
    qtys = defaultdict(list)
    for r in rows:
        prices[r["impa"]].append(r["price"])
        qtys[r["impa"]].append(r["qty"] or 0)
    med_price = {k: median(v) for k, v in prices.items()}
    med_qty = {k: median(v) for k, v in qtys.items()}

    period_rows = [r for r in rows if r["quoted_at"] >= cutoff]

    outliers = []
    for r in period_rows:
        impa = r["impa"]
        mp = med_price.get(impa)
        mq = med_qty.get(impa)
        reasons = []
        if len(prices[impa]) == 1:
            reasons.append("first-time item")
        elif mp and abs(r["price"] - mp) / mp > 0.20:
            reasons.append(f"price {r['price']:.2f} vs usual {mp:.2f}")
        if mq and r["qty"] and r["qty"] > 3 * mq:
            reasons.append(f"qty {r['qty']:.0f} vs usual {mq:.0f}")
        if reasons:
            outliers.append((r, "; ".join(reasons)))

    stats = defaultdict(lambda: {"n": 0, "qty": 0.0, "prices": [], "desc": "", "brand": "", "last": ""})
    for r in period_rows:
        s = stats[r["impa"]]
        s["n"] += 1
        s["qty"] += (r["qty"] or 0)
        s["prices"].append(r["price"])
        s["desc"] = r["description"] or s["desc"]
        s["brand"] = r["brand"] or s["brand"]
        s["last"] = r["quoted_at"]

    total_value = sum(r["price"] * (r["qty"] or 0) for r in period_rows)

    if output_path is None:
        os.makedirs(REPORTS_DIR, exist_ok=True)
        output_path = os.path.join(REPORTS_DIR, f"report_{now:%Y-%m-%d}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Monthly quote report", now.strftime("%Y-%m-%d")])
    ws.append([])
    ws.append(["Period (days)", period_days])
    ws.append(["Lines quoted", len(period_rows)])
    ws.append(["Distinct items", len(stats)])
    ws.append(["Total quoted value", round(total_value, 2)])
    ws.append(["Outliers flagged", len(outliers)])

    ws2 = wb.create_sheet("Items")
    ws2.append(["IMPA", "Description", "Times quoted", "Total qty", "Min price", "Avg price", "Max price", "Brand", "Last quoted"])
    for impa, s in sorted(stats.items(), key=lambda x: -x[1]["n"]):
        ps = s["prices"]
        ws2.append([impa, s["desc"], s["n"], s["qty"],
                    round(min(ps), 2), round(sum(ps) / len(ps), 2), round(max(ps), 2),
                    s["brand"], s["last"]])

    ws3 = wb.create_sheet("Outliers")
    ws3.append(["Date", "RFQ", "IMPA", "Description", "Qty", "Price", "Reason"])
    for r, reason in sorted(outliers, key=lambda x: x[0]["quoted_at"]):
        ws3.append([r["quoted_at"], r["rfq_ref"], r["impa"], r["description"], r["qty"], r["price"], reason])

    wb.save(output_path)
    return output_path

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "report":
        p = generate_report()
        if p:
            print("Report saved:", p)
    else:
        print("Usage: python tracker.py report")
