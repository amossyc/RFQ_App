import re
import os
import glob
import sys
import win32com.client
import config

RFQ_ENTRY_RE = re.compile(
    r"https?://" + re.escape(config.PORTAL_DOMAIN or "your-portal.example.com") + r"/portal-ext/entry\?access=[A-Za-z0-9]+",
    re.IGNORECASE,
)
RFQ_SENDER = config.SENDER_EMAIL

def get_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    return ns.GetDefaultFolder(6)  # 6 = Inbox

def extract_entry_link(text):
    m = RFQ_ENTRY_RE.search(text or "")
    return m.group(0) if m else None

def reply_done(msg, body="", attachment=None, link=None):
    reply = msg.ReplyAll()  # reply to sender + all other recipients/CC
    original = reply.Body  # quoted original message from the sender
    parts = []
    if link:
        parts.append(link)  # portal link pinned to the very top for easy clicking
    if body:
        parts.append(body)
    reply.Body = "\n\n".join(parts) + (("\n\n" + original) if original else "")
    if attachment:
        reply.Attachments.Add(attachment)
    reply.Send()
    return True

PROCESSED_FOLDER = "SINWA RFQ Processed"
FAILED_FOLDER = "SINWA RFQ Failed"

def _get_subfolder(name):
    inbox = get_inbox()
    try:
        return inbox.Folders[name]
    except Exception:
        return inbox.Folders.Add(name)

def get_processed_folder():
    return _get_subfolder(PROCESSED_FOLDER)

def get_failed_folder():
    return _get_subfolder(FAILED_FOLDER)

def move_to_folder(msg, folder):
    """Move an email into the given Outlook folder."""
    try:
        msg.Move(folder)
        return True
    except Exception:
        return False

def move_to_processed(msg):
    """Move a finished RFQ email out of the Inbox into the processed folder."""
    return move_to_folder(msg, get_processed_folder())

def move_to_failed(msg):
    """Move an RFQ email that failed import 3 times into the failed folder."""
    return move_to_folder(msg, get_failed_folder())

PRICELIST_FOLDER = "SINWA Pricelist History"

def _scan_for_attachment(match):
    """Find inbox emails that have an attachment whose filename satisfies match(name)."""
    inbox = get_inbox()
    items = inbox.Items
    try:
        items = items.Restrict("[MessageClass]='IPM.Note'")
    except Exception:
        pass
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    result = []
    for msg in items:
        try:
            if any(match(str(a.FileName)) for a in msg.Attachments):
                result.append({
                    "msg": msg,
                    "subject": msg.Subject,
                    "entry_id": getattr(msg, "EntryID", None),
                })
        except Exception:
            continue
    return result

def _is_pricelist_name(name):
    n = name.lower()
    return n.endswith(".xlsm") and "sinwa qtn" in n

def _is_no_offer_name(name):
    n = name.lower()
    return n.endswith(".xlsx") and ("no_offer" in n or "no offer" in n or "nooffer" in n)

def get_pricelist_emails():
    """Inbox emails carrying a pricelist (.xlsm with 'sinwa qtn' in the name)."""
    return _scan_for_attachment(_is_pricelist_name)

def get_no_offer_emails():
    """Inbox emails carrying the no-offer list (.xlsx with 'no offer' in the name)."""
    return _scan_for_attachment(_is_no_offer_name)

def save_pricelist_attachment(msg, dest_dir):
    """Save the pricelist .xlsm attachment into dest_dir; remove old pricelists first."""
    for att in msg.Attachments:
        name = str(att.FileName)
        if _is_pricelist_name(name):
            # old pricelist not kept: remove any existing 'sinwa qtn' .xlsm files
            for old in glob.glob(os.path.join(dest_dir, "*.xlsm")):
                b = os.path.basename(old)
                if "sinwa qtn" in b.lower() and not b.startswith("~$"):
                    try:
                        os.remove(old)
                    except Exception:
                        pass
            path = os.path.join(dest_dir, name)
            att.SaveAsFile(path)
            return path
    return None

def save_no_offer_attachment(msg, dest_dir):
    """Save the no-offer .xlsx attachment as no_offer.xlsx into dest_dir (overwrite)."""
    for att in msg.Attachments:
        name = str(att.FileName)
        if _is_no_offer_name(name):
            path = os.path.join(dest_dir, "no_offer.xlsx")
            att.SaveAsFile(path)
            return path
    return None

def process_pricelist_emails(dest_dir):
    """Save pricelist attachments into dest_dir and move those emails to the history folder."""
    saved = []
    for e in get_pricelist_emails():
        try:
            path = save_pricelist_attachment(e["msg"], dest_dir)
            if path:
                saved.append(path)
                move_to_folder(e["msg"], _get_subfolder(PRICELIST_FOLDER))
        except Exception as ex:
            print(f"  pricelist save failed ({e['subject']}): {ex}")
    return saved

def process_no_offer_emails(dest_dir):
    """Save the no-offer list into dest_dir and move those emails to the history folder."""
    saved = []
    for e in get_no_offer_emails():
        try:
            path = save_no_offer_attachment(e["msg"], dest_dir)
            if path:
                saved.append(path)
                move_to_folder(e["msg"], _get_subfolder(PRICELIST_FOLDER))
        except Exception as ex:
            print(f"  no_offer save failed ({e['subject']}): {ex}")
    return saved

def _scan_folder(folder, unread_only=False):
    items = folder.Items
    try:
        items = items.Restrict("[MessageClass]='IPM.Note'")
    except Exception:
        pass
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    result = []
    for msg in items:
        try:
            link = extract_entry_link(msg.Body) or extract_entry_link(msg.HTMLBody)
            if not link:
                continue
            if unread_only and not msg.Unread:
                continue
            result.append({
                "msg": msg,
                "link": link,
                "subject": msg.Subject,
                "entry_id": getattr(msg, "EntryID", None),
            })
        except Exception:
            continue
    return result

def get_rfq_emails(unread_only=False, include_processed=False):
    result = _scan_folder(get_inbox(), unread_only)
    if include_processed:
        for getter in (get_processed_folder, get_failed_folder):
            try:
                result.extend(_scan_folder(getter(), unread_only))
            except Exception:
                pass
    return result

def get_email_table_rfqs():
    """Inbox emails that carry an RFQ table in the body but no portal entry link."""
    inbox = get_inbox()
    items = inbox.Items
    try:
        items = items.Restrict("[MessageClass]='IPM.Note'")
    except Exception:
        pass
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    result = []
    for msg in items:
        try:
            link = extract_entry_link(msg.Body) or extract_entry_link(msg.HTMLBody)
            if link:
                continue
            if "<table" not in (msg.HTMLBody or "").lower():
                continue
            result.append({
                "msg": msg,
                "subject": msg.Subject,
                "entry_id": getattr(msg, "EntryID", None),
            })
        except Exception:
            continue
    return result

def scan(n=50):
    inbox = get_inbox()
    items = inbox.Items
    try:
        items = items.Restrict("[MessageClass]='IPM.Note'")
    except Exception:
        pass
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    shown = 0
    for msg in items:
        if shown >= n:
            break
        try:
            link = extract_entry_link(msg.Body) or extract_entry_link(msg.HTMLBody)
            unread = "UNREAD" if msg.Unread else "read  "
            print(f"{shown+1:>2}. [{unread}] {str(msg.ReceivedTime)[:16]} | {msg.SenderEmailAddress}")
            print(f"     subj: {msg.Subject}")
            if link:
                print(f"     ENTRY LINK: {link}")
            shown += 1
        except Exception as e:
            shown += 1
            print(f"{shown:>2}. (skipped an item: {e})")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "scan":
        scan()
    else:
        print("Usage: python outlook.py scan")
