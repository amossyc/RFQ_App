import os
import sys
from playwright.sync_api import sync_playwright
import config

PROFILE = config.PROFILE
WORK = config.WORK

def _launch():
    p = sync_playwright().start()
    ctx = p.chromium.launch_persistent_context(
        PROFILE,
        channel="msedge",
        headless=False,
        accept_downloads=True,
        args=["--no-first-run", "--no-default-browser-check", "--disable-features=msEdgeFirstRunExperience"],
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    return p, ctx, page

def export_rfq(url, dest_dir=WORK):
    p, ctx, page = _launch()
    try:
        print("Opening RFQ page ...")
        page.goto(url, wait_until="domcontentloaded")
        btn = page.get_by_text("Export", exact=True).first
        btn.wait_for(state="visible", timeout=30000)
        print("Clicking Export ...")
        with page.expect_download(timeout=60000) as dl:
            btn.click()
        download = dl.value
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, download.suggested_filename)
        download.save_as(dest)
        print("Downloaded:", dest)
        return dest
    finally:
        try:
            ctx.close()
        except Exception:
            pass
        try:
            p.stop()
        except Exception:
            pass

def import_quote(url, file_path):
    p, ctx, page = _launch()
    try:
        print("Opening RFQ page ...")
        page.goto(url, wait_until="domcontentloaded")
        btn = page.get_by_text("Import", exact=True).first
        btn.wait_for(state="visible", timeout=30000)
        print("Clicking Import and selecting file ...")
        with page.expect_file_chooser(timeout=30000) as fc:
            btn.click()
        fc.value.set_files(file_path)
        print("File selected. Waiting for upload ...")
        page.wait_for_timeout(5000)
        print("Import action complete (Submit was NOT clicked).")
        return True
    finally:
        try:
            ctx.close()
        except Exception:
            pass
        try:
            p.stop()
        except Exception:
            pass

def process_rfq(entry_url, fill_callback, dest_dir=WORK):
    """Open the token entry link, export, fill, and import — all in ONE session
    (so the one-time token is only used once). fill_callback(exported_path)
    must return the filled file path."""
    p, ctx, page = _launch()
    try:
        print("Opening entry link ...")
        page.goto(entry_url, wait_until="domcontentloaded")
        export_btn = page.get_by_text("Export", exact=True).first
        export_btn.wait_for(state="visible", timeout=30000)
        print("Landed on:", page.url)

        print("Clicking Export ...")
        with page.expect_download(timeout=60000) as dl:
            export_btn.click()
        download = dl.value
        os.makedirs(dest_dir, exist_ok=True)
        exported = os.path.join(dest_dir, download.suggested_filename)
        download.save_as(exported)
        print("Downloaded:", exported)

        print("Filling ...")
        filled = fill_callback(exported)
        print("Filled:", filled)

        print("Clicking Import ...")
        import_btn = page.get_by_text("Import", exact=True).first
        import_btn.wait_for(state="visible", timeout=30000)
        with page.expect_file_chooser(timeout=30000) as fc:
            import_btn.click()
        fc.value.set_files(filled)
        page.wait_for_timeout(5000)
        print("Import complete (Submit NOT clicked).")
        return exported, filled
    finally:
        try:
            ctx.close()
        except Exception:
            pass
        try:
            p.stop()
        except Exception:
            pass

if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "export":
        export_rfq(sys.argv[2])
    elif len(sys.argv) >= 4 and sys.argv[1] == "import":
        import_quote(sys.argv[2], sys.argv[3])
    else:
        print("Usage:")
        print('  python portal.py export "URL"')
        print('  python portal.py import "URL" "C:\\path\\file.xlsx"')
