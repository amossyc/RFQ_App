import os

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, "config.txt")

def _resolve(v):
    v = v.strip().strip('"').strip("'")
    if not v:
        return None
    if os.path.isabs(v):
        return v
    return os.path.join(APP_DIR, v)

_defaults = {
    "pricelist_dir": os.path.join(APP_DIR, "pricelist"),
    "working_dir": os.path.join(APP_DIR, "working"),
    "browser_profile": os.path.join(APP_DIR, "browser_profile"),
    "reports_dir": os.path.join(APP_DIR, "reports"),
    "db": os.path.join(APP_DIR, "tracking.db"),
    "check_interval_minutes": "15",
    "sender_email": "",
    "sheet_name": "",
    "portal_domain": "",
}

_PATH_KEYS = {"pricelist_dir", "working_dir", "browser_profile", "reports_dir", "db"}

def _load():
    cfg = dict(_defaults)
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip().lower()
                v = v.strip().strip('"').strip("'")
                if not v:
                    continue
                cfg[k] = _resolve(v) if k in _PATH_KEYS else v
    return cfg

_cfg = _load()

PRICELIST_DIR = _cfg["pricelist_dir"]
WORKING_DIR = _cfg["working_dir"]
BROWSER_PROFILE = _cfg["browser_profile"]
REPORTS_DIR = _cfg["reports_dir"]
DB = _cfg["db"]

# aliases used by existing scripts
WORK = WORKING_DIR
PROFILE = BROWSER_PROFILE

CHECK_INTERVAL_MINUTES = int(_cfg.get("check_interval_minutes", "15"))

SENDER_EMAIL = _cfg.get("sender_email", "")
SHEET_NAME = _cfg.get("sheet_name", "")
PORTAL_DOMAIN = _cfg.get("portal_domain", "")
