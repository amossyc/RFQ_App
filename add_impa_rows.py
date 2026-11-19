import glob, os, shutil, datetime, sys
import openpyxl
import config

# usage: python add_impa_rows.py <START> <END>
if len(sys.argv) >= 3:
    START, END = int(sys.argv[1]), int(sys.argv[2])
else:
    START, END = 754405, 755164

# latest pricelist, ignoring Excel lock files (~$...)
files = [f for f in glob.glob(os.path.join(config.PRICELIST_DIR, "*.xlsm"))
         if not os.path.basename(f).startswith("~$")]
if not files:
    raise SystemExit("ERROR: no pricelist found")
pl = sorted(files, key=os.path.getmtime)[-1]
print("Pricelist:", pl)

# 1) backup
backup = pl + ".bak_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
shutil.copy2(pl, backup)
print("Backup saved to:", backup)

# 2) load (keep VBA macros)
wb = openpyxl.load_workbook(pl, keep_vba=True)
ws = wb[config.SHEET_NAME or wb.sheetnames[0]]

# find the "Net Price" column from the header (supports the ALT layout)
header = [str(ws.cell(row=1, column=c).value or "").strip().lower() for c in range(1, 30)]
net_col = None
for i, h in enumerate(header):
    if h == "net price":
        net_col = i + 1  # 1-based
        break
if net_col is None:
    net_col = 9  # fallback: old layout
print("Net Price column:", net_col)

start_row = end_row = None
for r in range(2, ws.max_row + 1):
    v = ws.cell(row=r, column=1).value
    if v == START:
        start_row = r
    elif v == END:
        end_row = r
        break

if not start_row or not end_row:
    raise SystemExit(f"ERROR: could not find {START} or {END} in column A")

print(f"{START} at row {start_row}, {END} at row {end_row}")

full = list(range(START + 1, END))          # every code that should sit between them
existing_count = end_row - start_row - 1    # rows already there
print(f"Existing rows between them: {existing_count} | full range: {len(full)}")

# preserve existing rows' full data (description/brands) so nothing is lost
existing_data = {}
for r in range(start_row + 1, end_row):
    code = ws.cell(row=r, column=1).value
    vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
    existing_data[str(code).strip()] = vals

missing = len(full) - existing_count
if missing < 0:
    raise SystemExit(f"ERROR: {existing_count} rows present but only {len(full)} codes in range")

if missing > 0:
    ws.insert_rows(start_row + 1, amount=missing)
    print(f"Inserted {missing} blank rows.")

# 3) write the whole block in sorted order; restore existing rows, fill gaps with IMPA + Net 0
for i, code in enumerate(full):
    r = start_row + 1 + i
    key = str(code).strip()
    if key in existing_data:
        for c, val in enumerate(existing_data[key], start=1):
            ws.cell(row=r, column=c).value = val
    else:
        ws.cell(row=r, column=1).value = code         # column A = IMPA (int)
        ws.cell(row=r, column=net_col).value = 0      # Net Price = 0

wb.save(pl)
print("Saved:", pl)

# 4) verify
wb2 = openpyxl.load_workbook(pl, data_only=True, read_only=True)
ws2 = wb2[config.SHEET_NAME or wb2.sheetnames[0]]
for r in (start_row, start_row + 1, start_row + len(full), start_row + len(full) + 1):
    print(f"  row {r}: A={ws2.cell(row=r, column=1).value!r}  Net={ws2.cell(row=r, column=net_col).value!r}")

print(f"\nDONE. Block now has {len(full)} codes ({full[0]} .. {full[-1]}).")
