"""One-off: re-process an email-list RFQ and reply.

Usage:
  python retry_email_rfq.py <SUBJECT>             -> Reply-All
  python retry_email_rfq.py <SUBJECT> <TO_ADDRESS> -> reply to that address only
"""
import sys, os, re
import win32com.client
import fill_quote
import email_rfq
import main
import outlook


def find_email(subj):
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(6)
    for msg in inbox.Items:
        if subj.upper() in (msg.Subject or "").upper():
            return msg
    return None


if __name__ == "__main__":
    subj = sys.argv[1] if len(sys.argv) > 1 else "TQ26174179"
    to = sys.argv[2] if len(sys.argv) > 2 else None

    msg = find_email(subj)
    if not msg:
        print("email not found:", subj)
        raise SystemExit(1)

    html = msg.HTMLBody or ""
    # use the production pricelist folder (the workspace config.txt points at Downloads for testing)
    pl = fill_quote.newest(r"C:\RFQ_App\pricelist\*.xlsm")
    ref = main.extract_ref(msg.Subject) or main._safe_filename(main.extract_tq(msg.Subject))
    out = os.path.join(r"C:\RFQ_App\working", f"{ref}-done.xlsx")

    rep = email_rfq.process_email_rfq(html, out, pricelist_path=pl)
    if not rep:
        print("no table found in the email")
        raise SystemExit(1)

    summary = main.build_summary(ref, rep)
    html_table = rep.get("html_table")

    # save the modified table as an HTML file for easy copy/paste
    if html_table:
        html_out = out.replace(".xlsx", "-table.html")
        with open(html_out, "w", encoding="utf-8") as f:
            f.write("<html><head><meta charset='utf-8'></head><body>" + html_table + "</body></html>")
        print("HTML table saved:", html_out)

    if to:
        reply = msg.Reply()      # reply to sender, then override recipient
        reply.To = to
    else:
        reply = msg.ReplyAll()   # reply-all by default
    outlook.strip_image_attachments(reply)  # drop inline signature/logos
    reply.Body = summary + "\n\n" + (reply.Body or "")
    if html_table:
        # prepend the clean completed table, keep the customer's original table below
        orig_html = reply.HTMLBody or ""
        if "<body" in orig_html.lower():
            reply.HTMLBody = re.sub(r"(<body[^>]*>)", r"\1" + html_table + "<br><br><hr><br>", orig_html, count=1, flags=re.IGNORECASE)
        else:
            reply.HTMLBody = "<html><body>" + html_table + "<br><br>" + orig_html + "</body></html>"
        reply.BodyFormat = 2  # olFormatHTML — force HTML so tables aren't flattened
    reply.Attachments.Add(out)
    reply.Send()
    print(f"Replied {'to ' + to if to else 'all'}")
    print(f"Attachment: {out}")
