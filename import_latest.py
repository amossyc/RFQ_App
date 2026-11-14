import sys
import os
import glob
import portal
import config

WORK = config.WORK

def newest_filled():
    files = glob.glob(os.path.join(WORK, "*-done.xlsx"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('Usage: python import_latest.py "https://<portal>/portal-ext/rfq/XXXXX"')
        sys.exit(1)
    url = sys.argv[1]
    f = newest_filled()
    if not f:
        print("No *-done.xlsx found in", WORK)
        print("Run 'python fill_quote.py' first.")
        sys.exit(1)
    print("Importing:", os.path.basename(f))
    portal.import_quote(url, f)
