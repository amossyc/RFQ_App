import sys
import os
import json
import re
import outlook
import portal
import fill_quote
import tracker

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PROCESSED_FILE = os.path.join(APP_DIR, "processed.json")

def extract_tq(subject):
    m = re.search(r"TQ\d+", subject or "")
    return m.group(0) if m else (subject or "")

def extract_ref(subject):
    """Return the RFQ reference (TRFQ preferred, else TQ) for naming output files."""
    m = re.search(r"TRFQ\d+", subject or "")
    if m:
        return m.group(0)
    m = re.search(r"TQ\d+", subject or "")
    return m.group(0) if m else None

def load_processed():
    if os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE, encoding="utf-8") as f:
            return set(json.load(f))
    return set()

def save_processed(s):
    with open(PROCESSED_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(s), f)

def make_fill(prclist, ref=None):
    holder = {}
    def fill(exported):
        if ref:
            out_name = f"{ref}-done.xlsx"
        else:
            out_name = os.path.basename(exported).replace(".xlsx", "-done.xlsx")
        out = os.path.join(os.path.dirname(exported), out_name)
        rep = fill_quote.fill_quote(exported, prclist, out)
        holder["report"] = rep
        print(f"   fill: {rep['total']} items, {len(rep['flagged'])} flagged, {len(rep.get('auto_matched', []))} auto-matched")
        for f in rep["flagged"]:
            print("      FLAG (manual):", f[0], f[1], f[2][:45])
        for a in rep.get("auto_matched", []):
            print("      AUTO-MATCH (review):", a[0], "->", a[1], a[2][:45])
        return out
    return fill, holder

def build_summary(rfq_ref, rep):
    flagged = rep.get("flagged", [])
    auto = rep.get("auto_matched", [])
    quoted = rep.get("total", 0) - len(flagged)
    lines = [f"RFQ {rfq_ref} processed."]
    lines.append(f"Quoted: {quoted} lines | Flagged: {len(flagged)} | Auto-matched: {len(auto)}")
    if auto:
        lines.append("")
        lines.append("Auto-matched (verify):")
        for a in auto:
            lines.append(f"  - line {a[0]}: {a[2][:40]} -> IMPA {a[1]}")
    mismatch = [f for f in flagged if "[IMPA mismatch]" in f[2]]
    other = [f for f in flagged if "[IMPA mismatch]" not in f[2]]
    if mismatch:
        lines.append("")
        lines.append("IMPA mismatch (wrong code - please verify):")
        for f in mismatch:
            name = f[2].replace("[IMPA mismatch] ", "", 1)
            lines.append(f"  - line {f[0]}: {name[:40]}  (code {f[1]})")
    if other:
        lines.append("")
        lines.append("Flagged (need manual quote):")
        for f in other:
            lines.append(f"  - line {f[0]}: {f[2][:40]}")
    return "\n".join(lines)

def run(dry_run=False, limit=None, tq_filters=None, skip_filters=None, force_filters=None, send_reply=True):
    processed = load_processed()
    prclist = fill_quote.newest(os.path.join(fill_quote.PRICELIST_DIR, "*.xlsm"))
    if not prclist:
        print("No pricelist found in", fill_quote.PRICELIST_DIR)
        return
    print("Pricelist:", prclist)

    emails = outlook.get_rfq_emails()
    if force_filters:
        emails = [e for e in emails if any(f in (e["subject"] or "").upper() for f in force_filters)]
    if tq_filters:
        emails = [e for e in emails if any(tq in (e["subject"] or "").upper() for tq in tq_filters)]
    if skip_filters:
        emails = [e for e in emails if not any(sk in (e["subject"] or "").upper() for sk in skip_filters)]
    print("Found", len(emails), "RFQ email(s).")

    done = 0
    for e in emails:
        key = e.get("entry_id") or e["subject"]
        is_forced = bool(force_filters) and any(f in (e["subject"] or "").upper() for f in force_filters)
        if key in processed and not is_forced:
            print("skip (already processed):", e["subject"])
            continue
        print("=" * 62)
        print("RFQ:", e["subject"])
        print("  link:", e["link"])
        if dry_run:
            print("  [DRY RUN] would export -> fill -> import -> reply (with summary)")
            continue
        try:
            ref = extract_ref(e["subject"])
            fill_fn, holder = make_fill(prclist, ref)
            exported, filled = portal.process_rfq(e["link"], fill_fn)
            if exported and os.path.exists(exported):
                try:
                    os.remove(exported)  # keep only the -done version
                except Exception:
                    pass
            rep = holder.get("report", {})
            lines = tracker.read_filled_lines(filled)
            tracker.log_quote(tracker.DB, extract_tq(e["subject"]), e["subject"], lines)
            summary = build_summary(extract_tq(e["subject"]), rep)
            if send_reply:
                outlook.reply_done(e["msg"], summary)
            processed.add(key)
            save_processed(processed)
            status = "replied with summary" if send_reply else "no reply sent"
            print(f"  COMPLETE ({status}, logged {len(lines)} lines)")
        except portal.ImportFailed as ex:
            print("  IMPORT FAILED:", ex)
            if send_reply:
                outlook.reply_done(e["msg"], "Import failed")
            processed.add(key)
            save_processed(processed)
        except Exception as ex:
            print("  ERROR:", ex)
        done += 1
        if limit and done >= limit:
            break
    print("\nFinished. Newly processed:", done)

def process_direct_link(link):
    """Process a direct RFQ link (export -> fill -> import -> log), no email/reply."""
    prclist = fill_quote.newest(os.path.join(fill_quote.PRICELIST_DIR, "*.xlsm"))
    if not prclist:
        print("No pricelist found in", fill_quote.PRICELIST_DIR)
        return
    print("Pricelist:", prclist)
    print("Processing direct link:", link)
    ref = "DIRECT_" + link.rstrip("/").split("/")[-1]
    fill_fn, holder = make_fill(prclist, ref)
    try:
        exported, filled = portal.process_rfq(link, fill_fn)
        if exported and os.path.exists(exported):
            try:
                os.remove(exported)  # keep only the -done version
            except Exception:
                pass
        rep = holder.get("report", {})
        lines = tracker.read_filled_lines(filled)
        ref = "DIRECT-" + link.rstrip("/").split("/")[-1]
        tracker.log_quote(tracker.DB, ref, link, lines)
        print(f"  COMPLETE (no reply, logged {len(lines)} lines)")
        print(f"  fill: {rep['total']} items, {len(rep['flagged'])} flagged, {len(rep.get('auto_matched', []))} auto-matched")
        for f in rep["flagged"]:
            print("      FLAG (manual):", f[0], f[1], f[2][:45])
        for a in rep.get("auto_matched", []):
            print("      AUTO-MATCH (review):", a[0], "->", a[1], a[2][:45])
    except portal.ImportFailed as ex:
        print("  IMPORT FAILED:", ex)
    except Exception as ex:
        print("  ERROR:", ex)

if __name__ == "__main__":
    if "--link" in sys.argv:
        link = sys.argv[sys.argv.index("--link") + 1]
        process_direct_link(link)
        sys.exit()
    dry = "--dry-run" in sys.argv
    limit = None
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])
    tq = None
    if "--tq" in sys.argv:
        vals = [x.strip().upper() for x in sys.argv[sys.argv.index("--tq") + 1].split(",") if x.strip()]
        tq = vals or None
    skip = None
    if "--skip" in sys.argv:
        skip = [x.strip().upper() for x in sys.argv[sys.argv.index("--skip") + 1].split(",") if x.strip()] or None
    force = None
    if "--force" in sys.argv:
        force = [x.strip().upper() for x in sys.argv[sys.argv.index("--force") + 1].split(",") if x.strip()] or None
    no_reply = "--no-reply" in sys.argv
    run(dry_run=dry, limit=limit, tq_filters=tq, skip_filters=skip, force_filters=force, send_reply=not no_reply)
