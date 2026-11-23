# RFQ Automation App

Automates the end-to-end RFQ (request-for-quotation) workflow for the customer's procurement portal:
read RFQ emails from Outlook → export the RFQ from the portal → fill prices/remarks from
a pricelist → import the completed file back → reply to the email with a summary → track
history and produce a monthly outlier report.

It is a set of plain Python scripts (no framework), designed to run on Windows against
Microsoft Outlook (via `pywin32`) and Microsoft Edge (via Playwright).

---

## 1. End-to-end workflow (one RFQ)

1. `main.py` scans the Outlook Inbox for emails that contain an `/entry?access=<token>`
   link (RFQ emails are identified by the link, not the sender address).
2. For each unprocessed email, `portal.process_rfq()` opens the link in a persistent Edge
   profile, clicks **Export** (downloads a randomly-named `.xlsx`), runs the fill, then
   clicks **Import** to upload the filled file. It never clicks **Submit** — a human
   reviews and submits on the portal. If **Import** fails (button not clickable within
   10 seconds), the email is retried on later runs, up to **3 attempts**; after the 3rd
   failure it is moved to the `SINWA RFQ Failed` folder and replied to with `Import failed`.
3. `fill_quote.py` fills each line's List Price and Remarks from the pricelist.
4. The filled file is saved as `<TRFQ>-done.xlsx` and the empty export is deleted.
5. `outlook.reply_done()` replies (Reply-All) with a short summary (quoted/flagged/
   auto-matched counts + details), keeping the original message body, and attaches the
   filled `<TRFQ>-done.xlsx`.
6. `tracker.log_quote()` records every filled line in SQLite (`tracking.db`).
7. The email is moved to the `SINWA RFQ Processed` folder, and its `EntryID` is added to
   `processed.json` so it is not processed again.

The app is idempotent: re-running only processes *new* emails. `--force <TRFQ>` re-runs a
specific RFQ regardless.

---

## 2. Setup

Requirements (Windows):

- Python 3.8+
- `pip install playwright pywin32 openpyxl`
- `playwright install` (uses the system **Microsoft Edge** channel; no separate browser
  download needed)
- Microsoft Outlook (desktop) configured with the RFQ mailbox.

First run:

1. Put the pricelist `.xlsm` into `pricelist/` (set the sheet name via `sheet_name` in
   `config.txt`). Optionally put a `no_offer.xlsx` there too (IMPA codes to never quote).
2. Edit `config.txt`:
   - set `sheet_name` to the pricelist worksheet name,
   - set `portal_domain` to the portal host,
   - optionally set `check_interval_minutes` and `impa_check`.
   (`sender_email` is kept for reference only; RFQ emails are identified by their entry
   link, not the sender.)
3. Run `python main.py --dry-run` to verify it can find emails and the pricelist without
   touching anything.
4. Run `python main.py` (or double-click `1_Process Inbox.bat`) to process for real.
5. The first portal login happens in the persistent browser profile
   (`browser_profile/`); after that it stays logged in.

Pricelist / no-offer updates can also arrive by email: each scheduler cycle checks the
inbox for a `.xlsm` attachment whose name contains `sinwa qtn` (pricelist) or a `.xlsx`
attachment whose name contains `no offer` (no-offer list), saves it into `pricelist/`, and
moves the email to the `SINWA Pricelist History` folder.

Double-click shortcuts (all run from their own folder via `%~dp0`):

| File | What it does |
|---|---|
| `1_Process Inbox.bat` | process all new emails (sends replies) |
| `2_Start Scheduler.bat` | loop forever, check inbox every N minutes |
| `3_Process Single RFQ.bat` | prompts for one TRFQ, processes only it (sends reply) |
| `4_Process Single Link.bat` | prompts for a portal URL (for emails that arrived with no link) |
| `5_Process Table Converter.bat` | prompts for an email-list RFQ, re-quotes it and replies (Reply-All by default) |
| `6_Monthly Report.bat` | generate the monthly outlier report |

---

## 3. File layout

| File | Role |
|---|---|
| `main.py` | Orchestrator / CLI. Email loop, reply, logging, idempotency. |
| `portal.py` | Playwright browser automation (export / import / process_rfq). |
| `fill_quote.py` | **The core business logic.** Pricelist parsing, description matching, quote building. |
| `tracker.py` | SQLite storage + monthly Excel outlier report. |
| `outlook.py` | Outlook COM: inbox scan, link extraction, reply. |
| `config.py` | Reads `config.txt` (paths + settings). |
| `scheduler.py` | Infinite loop: checks for pricelist/no-offer emails, then runs `main.run()` every N minutes. |
| `import_latest.py` | Small helper: re-import the newest `*-done.xlsx` into the portal. |

CLI (`python main.py`):

```
--dry-run              print what would happen, don't act
--limit N              stop after N emails
--tq A,B               only emails whose subject contains A or B
--skip X               skip emails whose subject contains X
--force Y              only re-process emails whose subject contains Y (used by Redo/Process Single)
--no-reply             don't send the reply email
--link <URL>           process one direct portal /rfq/<id> URL (no email)
```

---

## 4. Configuration (`config.txt`)

`key = value` lines; `#` comments; blank lines ignored. Path values are resolved against
the app folder if relative.

| Key | Meaning |
|---|---|
| `pricelist_dir` | folder containing the newest `*.xlsm` pricelist (and `no_offer.xlsx`) |
| `working_dir` | exports + completed `-done.xlsx` files |
| `browser_profile` | persistent Edge profile (keeps the portal logged in) |
| `reports_dir` | monthly report output |
| `db` | SQLite path (`tracking.db`) |
| `check_interval_minutes` | scheduler interval |
| `sender_email` | legacy/reference (RFQ emails are identified by entry link) |
| `sheet_name` | pricelist worksheet name (e.g. `SINWA`) |
| `portal_domain` | portal host (e.g. `sinwa.epbrowser.com`) |
| `impa_check` | `basic` (default) / `strict` / `off` — see §8 |

`config.py` exposes `PRICELIST_DIR`, `WORK` (alias), `PROFILE` (alias), `REPORTS_DIR`,
`DB`, `CHECK_INTERVAL_MINUTES`, `SENDER_EMAIL`, `SHEET_NAME`, `PORTAL_DOMAIN`, `IMPA_CHECK`.

---

## 5. Pricelist format (CRITICAL — read this first)

The pricelist is an `.xlsm`; the sheet name is set via `sheet_name` in `config.txt`. Columns
(detected from the header row, so a pricelist with or without the `ALT` column both work):

- `A` = IMPA code (6 digits). **May be blank** for items not yet in the IMPA catalog
  (those are still matched by description and stored under synthetic `NOCODE_n` keys).
- `B` = Stk ID, `C` = Doc Date, `D` = **ALT** (an alternative IMPA to quote instead).
- `E` = Description, `F` = UOM, `G` = date.
- `H` = Tender price, `I` = 2025 Tender, `J` = Net price.
- `K` onward = brand columns.

When a row's `ALT` cell points to another IMPA, that row quotes the target IMPA's
description/prices/brands (e.g. `851631` → `850741`).

### The "shifted" brand layout (non-obvious, important)

The brand columns are **not** simple `(brand, price)` pairs. They are laid out as:

```
K = PRIMARY brand          (no price of its own here; its price is in H/I/J)
L = price of the NEXT brand
M = next brand
N = price of the NEXT brand
O = next brand
... (price always precedes the brand it belongs to)
```

Example row:

```
K='EURO-CUT', L=2.05, M='DORMER UK', N=0.4, O='COBALT', P=5.6, Q='DORMER UK COBALT'
```

parses to `brands = [("EURO-CUT", None), ("DORMER UK", 2.05), ("COBALT", 0.4), ("DORMER UK COBALT", 5.6)]`.

`load_pricelist()` returns a dict: `IMPA -> {"tender": H, "tender25": I, "net": J, "brands": [...], "desc": E}`.
Rows with a blank `A` become keys `NOCODE_1`, `NOCODE_2`, … with `"no_code": True`.

`brands` is a list of `(name, price_or_None)`; the first element is the primary and its
price is `None` (price lives in the tender/net columns).

---

## 6. Quote building — the rules (`fill_quote.quote_item`)

`quote_item(entry, name, quality)` returns `(price, remark)`.

**Price selection** (`pick_price`): first non-empty of `tender25` (I) → `tender` (H) →
`net` (J).

**Cobalt filter**: unless `"COBALT"` appears in the customer description, brands containing
`COBALT` are dropped.

The remark logic runs in this order:

1. **AUTO DARK** (compact `AUTODARK` in the description) → pick the auto-dark brand(s):
   budget by default, budget + premium when quality is `good`.
2. **Stainless / non-spark** (`is_special`: `STAINLESS`, `S.STEEL`, `NON-SPARK` variants,
   or a standalone word `SS`) → always two options: `PRIMARY / $alt_price ALTERNATIVE`.
3. **Explicit brand named in the description** (`brand_in_description`):
   - named an *alternative* → quote **at that brand's own price**, with the primary as the
     cheap option: `NAMED / $primary_price PRIMARY`.
   - named the *primary* → quote at tender, with the first alternative: `PRIMARY / $alt ALT`.
   - (the primary case now uses the **full brand cell**, not a truncated word).
4. **Brand mentioned but not in the pricelist** (via `RefernenceNo : CODE,`) → standard
   two options.
5. **Quality tier**:
   - `good` → first `GOOD` brand (see sets below) **at its own price**; if none, first `AVG` brand.
   - `avg` → first `AVG` brand at its own price; if none, cheap + next-cheapest alternative.
6. **Default** → primary brand only.

Tier sets: `GOOD = {GERMANY, UK/GERMANY, BMI GERMANY, BAHCO, PYRAMEX, PRO-VISE, IMA ITALY}`,
`AVG = {FAPRO, STANLEY, TRAMONTINA, BOSCH, MAKITA}`.

`detect_quality` reads the RFQ's "Remarks:" field (kept as the **last** non-empty one, since
the template has a placeholder first) and/or the item description: `GOOD` → good;
`AVG`/`AVERAGE`/`FAPRO`/`IRENO` → avg.

### Special flags written into the sheet

- **IMPA mismatch** (see §8): List Price `0`, Remark `IMPA MISMATCH @@@@@@`.
- **`(CHK UOM)`** in a description → leave blank + flag (non-standard unit of measure).

---

## 7. Description matcher (`match_by_description`)

Used for lines with no code, or when the code lookup fails. Order of preference:

1. **AUTO DARK** feature (matches the item whose brands carry `AUTO DARK`).
2. **Priority features** (`NONSPARK`, `STAINLESS`, `100V`) — if the customer wrote one, the
   candidate must also have it (hard filter).
3. **Voltage** (`1000V`, `220V`, `110V`, …) — same hard filter.
4. **Size + word scoring**:
   - meaningful tokens (words) and size tokens (dimensions) are extracted;
   - score = shared words + 10 × shared size tokens, but the **size bonus only applies when
     ≥ 3 words already match** (so a coincidental size can't override the item type);
   - an item is only a candidate if `words + (1 if size matched else 0) >= 3`.
5. **Tie-breaks** (among equal top scores):
   - prefer `12-PT` > `6-PT`; `DEEP` only when the customer wrote `DEEP`;
   - then prefer the item whose leading words match the customer's wording.

Tokenizer normalisations (`_meaningful_tokens` / `_size_tokens`):

- `PHILLIPS` ↔ `PHILLIP` (same word).
- `NO.2` / `NO 2` → `NO2` (captures the Phillips/size number; `NO1`/`NO2`/`NO3` kept as tokens).
- sizes: `14MM`, `14 MM`, `14X17MM`, ranges `10-32MM`; `27X24` is normalised to `24X27`.
- decimal sizes stay one token (`2.0MM`, not `0MM`).
- `GENERIC_WORDS` (see the set in `fill_quote.py`) are ignored: includes `ELECTRIC`,
  `PNEUMATIC`, `HYDRAULIC`, `AUTOMATIC`, `MANUAL`, `POWER`, `INDUSTRIAL`, `DIA`, `HEAD`,
  `BODY`, `SPLIETHOFF`, `GROUP`, `CATALOG`, `REF`, and common English stopwords.
- The text is split at `| IMPA / | SSM / | Ref.` and at the metadata markers
  `Tender Category`, `Vessel Notes`, `Maker :` to drop catalog/metadata noise.

---

## 8. IMPA mismatch sanity check

When a line already has an IMPA code (from the portal's Stk column or the description), the
app verifies the code's pricelist description actually matches the customer's words. The
behaviour is controlled by `impa_check` in `config.txt`:

- **`basic`** (default): flag `IMPA mismatch` (price 0, remark `IMPA MISMATCH @@@@@@`) only
  when there is **zero** word overlap **and** zero size overlap, or a **voltage mismatch**
  (the customer specified a voltage the item does not have).
- **`strict`**: also flag weak matches (≥ 4 meaningful words but fewer than
  `min(4, pricelist keyword count)` shared keywords).
- **`off`**: never flag.

The `basic` mode avoids false positives from notation differences (`12.7MM` vs `1/2"`,
`SUS201` vs `STAINLESS STEEL`, marketing copy, `1,000V` vs `1000V`, etc.).

---

## 9. Tracking & monthly report

`tracking.db` (SQLite) logs every filled line: RFQ reference, line id, IMPA, name, qty,
price, remark. The monthly report (`python tracker.py report`) writes an Excel file to
`reports/` and highlights outliers:

- price > 20% off the median for the same IMPA,
- first-time items,
- quantity > 3× the median for the same IMPA.

---

## 10. Scheduler

`scheduler.py` loops forever: each cycle it (1) checks the inbox for pricelist / no-offer
emails and saves their attachments into `pricelist/`, (2) runs `main.run()` (process new
RFQ emails, send replies), then sleeps `check_interval_minutes` minutes and repeats. It
only stops on Ctrl+C (or closing the window). Errors are caught and logged so a transient
failure doesn't stop the loop.

- Pricelist email: `.xlsm` attachment with `sinwa qtn` in the filename → saved into
  `pricelist/` (any existing `sinwa qtn` `.xlsm` files are removed first) and the email
  moved to `SINWA Pricelist History`.
- No-offer email: `.xlsx` attachment with `no offer` in the filename → saved as
  `no_offer.xlsx` and the email moved to the same `SINWA Pricelist History` folder.

---

## 11. Idempotency

`processed.json` stores the set of processed email `EntryID`s. `main.run()` skips any email
already in it. `--force <TRFQ>` filters the inbox to *only* that RFQ and bypasses the
processed check (so it re-runs even if already done).

---

## 12. Known business rules / gotchas (history — don't re-derive these)

- **`999999`** is the customer's placeholder for "no IMPA code" — never treat it as a real code.
- `no_offer.xlsx` (IMPA codes to never quote) is a blacklist: any line resolving to one of
  those codes is left **fully blank** (no price, no remark) and not flagged.
- The customer **never buys** the non-pricelist items; flagging them is correct (do not add them
  to the pricelist automatically).
- "2025 tender price first, otherwise tender price, otherwise net price" is the price order.
- "first brand" is the default; description overrides: `GOOD`→good tier,
  `avg/average/fapro/ireno`→avg tier, `non-spark/stainless`→two options, explicit brand→that
  brand + a cheap option.
- A named **alternative** brand is quoted at **its own brand-column price** (not tender).
- `1000V` (and any voltage) is a **hard** requirement: quote a 1000V item when the customer
  says 1000V.
- `NO.2` / `#2` is the Phillips number and is a key differentiator for screwdrivers.
- Blank IMPA columns are intentional ("not in the IMPA catalog yet") and are still matched
  by description.
- The portal only reads List Price + Remarks on import; other columns are for the user's
  reference.
- The brand columns use the **shifted** layout (§5) — this has tripped up matching before.
- `TOWA` " / $" doubling was investigated and intentionally **left as-is** (a change was
  reverted).

---

## 13. Additional notes

- All the business logic lives in `fill_quote.py`; the other files are plumbing.
- The `-done.xlsx` files are filled by writing columns `I` (List Price), `K` (Remarks), and
  `L` (the parsed/matched IMPA, for reference — the Stk ID column `E` is left untouched).
  The RFQ line layout is: `C`=RFQ line id, `D`=Line number, `E`=Stk ID, `F`=Name,
  `G`=Qty, `H`=UOM, `I`=List Price, `J`=Line Total, `K`=Remarks, `L`=parsed IMPA.
- Voltages are normalised before matching: `1,000V` → `1000V` and `AC 220V` / `AC220V` →
  `220V`, so formatting differences don't cause a false voltage mismatch.
- There are many `test_*.py` scripts in the development workspace (not included here) that
  encode the expected behavior of the matcher/quote logic — worth re-creating if you change
  `fill_quote.py`.
- The portal is reached via the configured `portal_domain` (in `config.txt`); Export/Import are text elements (`Export`,
  `Import`); Import opens a hidden `<input type=file>` (Playwright `expect_file_chooser`).
- The pricelist changes frequently (daily/hourly); the app re-reads the `.xlsm` fresh on
  every run. Maintenance helper scripts live in the dev workspace: `add_impa_rows.py`
  (fill a range of missing IMPA codes with Net price 0) and `migrate_no_offer.py` (move
  net-0 placeholder rows from the pricelist into `no_offer.xlsx`).
