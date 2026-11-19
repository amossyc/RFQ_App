import glob, os, shutil, datetime, sys
import openpyxl
import config

# usage: python migrate_no_offer.py [--dry-run]
DRY = "--dry-run" in sys.argv

PL_DIR = config.PRICELIST_DIR
NO_OFFER = os.path.join(PL_DIR, "no_offer.xlsx")

files = [f for f in glob.glob(os.path.join(PL_DIR, "*.xlsm"))
         if not os.path.basename(f).startswith("~$")]
pl = sorted(files, key=os.path.getmtime)[-1]
print("Pricelist:", pl)
print("No-offer file:", NO_OFFER)

# 1) read all pricelist rows
wb = openpyxl.load_workbook(pl, data_only=True, read_only=True)
ws = wb[config.SHEET_NAME or wb.sheetnames[0]]
header_row = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
net_col = None
for i, h in enumerate(header_row, start=1):
    if str(h or "").strip().lower() == "net price":
        net_col = i
        break
if net_col is None:
    net_col = 10
print("Net Price column:", net_col)

keep = []       # rows to keep (full row tuples)
migrate = []    # IMPA codes to move to no_offer
for row in ws.iter_rows(min_row=2, values_only=True):
    impa = row[0]
    if impa in (None, ""):
        keep.append(row)  # no-code item -> keep
        continue
    net = row[net_col - 1] if (net_col - 1) < len(row) else None
    # any column other than IMPA (0) and Net Price (net_col-1) has info?
    has_other = any(v not in (None, "") for j, v in enumerate(row)
                    if j != 0 and j != (net_col - 1))
    if net == 0 and not has_other:
        migrate.append(int(str(impa).strip()))
    else:
        keep.append(row)
wb.close()

print(f"\nKeep (real items): {len(keep)}")
print(f"Migrate (IMPA + net 0 only): {len(migrate)}")

if DRY:
    print("\n[DRY RUN] no files changed.")
    if migrate:
        print("sample codes to migrate:", migrate[:15])
    raise SystemExit

# 2) write keep rows back to pricelist (preserve VBA)
backup = pl + ".bak_migrate_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
shutil.copy2(pl, backup)
print("\nBackup:", backup)

wb2 = openpyxl.load_workbook(pl, keep_vba=True)
ws2 = wb2[config.SHEET_NAME or wb2.sheetnames[0]]
if ws2.max_row > 1:
    ws2.delete_rows(2, ws2.max_row - 1)
for i, row in enumerate(keep):
    r = 2 + i
    for c, val in enumerate(row, start=1):
        ws2.cell(row=r, column=c).value = val
wb2.save(pl)
print(f"Pricelist shrunk to {len(keep)} items.")

# 3) rewrite no_offer.xlsx: existing + new codes, sorted, deduped
existing = set()
if os.path.exists(NO_OFFER):
    wb3 = openpyxl.load_workbook(NO_OFFER)
    ws3 = wb3.active
    for row in ws3.iter_rows(min_row=2, values_only=True):
        if row[0] not in (None, ""):
            try:
                existing.add(int(str(row[0]).strip()))
            except (ValueError, TypeError):
                pass
    wb3.close()

merged = sorted(existing | set(migrate))
new_added = len(set(migrate) - existing)
print(f"no_offer: existing={len(existing)} newly added={new_added} total={len(merged)}")

wb4 = openpyxl.Workbook()
ws4 = wb4.active
ws4.title = "Sheet1"
for i, h in enumerate(header_row[:11], start=1):  # header: IMPA .. BRAND
    ws4.cell(row=1, column=i).value = h
for i, code in enumerate(merged):
    r = 2 + i
    ws4.cell(row=r, column=1).value = code     # IMPA
    ws4.cell(row=r, column=10).value = 0       # Net Price
wb4.save(NO_OFFER)
print(f"no_offer.xlsx rewritten with {len(merged)} codes.")

print("\nDONE.")
