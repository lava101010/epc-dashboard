#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EPC Dashboard — Auto Updater (watch folder)
===========================================
Watches the  incoming/  folder for a new ERP export (0040_qa*.xls / .xlsx).
The export is CUMULATIVE (same sheet + new day rows) — that's fine: every run
rebuilds everything from the whole file.

When a NEW file appears (content hash differs from last processed):
   1) rebuild dist/  (website + standalone.html)
   2) regenerate EPC-Interactive-Dashboard.xlsx
   3) commit + push to GitHub (the Pages site refreshes automatically)
   4) move the file to incoming/processed/ and write tools/auto_update.log

Usage:
   python tools/auto_update.py            # single check (Task Scheduler friendly)
   python tools/auto_update.py --loop     # resident mode: check every 60 seconds

Recommended one-time setup (Windows, run once in cmd):
   schtasks /create /tn "EPC Auto Update" /tr "pythonw \"C:\EPC\tools\auto_update.py\" --loop" /sc onlogon
Then every day: export from ERP  ->  save/copy it into  incoming/  ->  done.
"""
import sys
import time
import json
import hashlib
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).resolve().parent.parent
INCOMING = ROOT / "incoming"
PROCESSED = INCOMING / "processed"
STATE = Path(__file__).resolve().parent / ".last_processed.json"
LOG = Path(__file__).resolve().parent / "auto_update.log"
INCOMING.mkdir(exist_ok=True)
PROCESSED.mkdir(exist_ok=True)


def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def newest_export():
    cands = []
    for pat in ("0040_qa*.xls", "0040_qa*.xlsx", "*.xls", "*.xlsx"):
        cands += [p for p in INCOMING.glob(pat) if not p.name.startswith("~$")]
    cands = [p for p in cands if p.parent == INCOMING]
    if not cands:
        return None
    return max(cands, key=lambda p: p.stat().st_mtime)


def load_state():
    try:
        return json.loads(STATE.read_text())
    except Exception:
        return {}


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True)
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def process(path):
    log(f"NEW export detected: {path.name} ({path.stat().st_size:,} bytes)")
    work = ROOT / "tools" / (".work" + path.suffix)
    shutil.copy(path, work)

    ok, out = run([sys.executable, str(ROOT / "tools" / "update_dashboard.py"), str(work)])
    if ok != 0:
        log("BUILD FAILED:\n" + out[-1500:])
        return False
    for line in out.splitlines():
        if any(k in line for k in ("Month ", "Produced", "Plan ", "Fingerprint")):
            log("   " + line.strip())

    ok2, out2 = run([sys.executable, str(ROOT / "tools" / "gen_excel.py"), str(work)])
    if ok2 == 0:
        log("   Excel workbook regenerated")
    else:
        log("   Excel regeneration FAILED:\n" + out2[-800:])

    okg, _ = run(["git", "-C", str(ROOT), "add", "-A"])
    okc, outc = run(["git", "-C", str(ROOT), "commit", "-q", "-m",
                     f"auto update {datetime.now():%Y-%m-%d %H:%M}"])
    if okc == 0:
        okp, outp = run(["git", "-C", str(ROOT), "push"])
        if okp == 0:
            log("   GitHub: pushed — Pages site refreshes in ~1 minute")
        else:
            log("   GitHub push failed (no remote or credentials?) — commit saved locally. " + outp.strip()[-200:])
    else:
        log("   GitHub: nothing new to commit")

    STATE.write_text(json.dumps({"hash": sha(path), "name": path.name,
                                 "time": str(datetime.now())}), encoding="utf-8")
    ts = datetime.now().strftime("%Y%m%d-%H%M")
    shutil.move(str(path), str(PROCESSED / f"{ts}-{path.name}"))
    try:
        work.unlink()
    except Exception:
        pass
    log("DONE — dashboard + Excel updated")
    return True


def check():
    p = newest_export()
    if p is None:
        return False
    h = sha(p)
    st = load_state()
    if st.get("hash") == h and st.get("name") == p.name:
        return False
    return process(p)


def main():
    if "--loop" in sys.argv:
        log("Auto updater STARTED (loop mode, every 60s).")
        log("Watch folder: " + str(INCOMING))
        while True:
            try:
                check()
            except Exception as e:
                log("ERROR: " + str(e))
            time.sleep(60)
    else:
        try:
            check()
        except Exception as e:
            log("ERROR: " + str(e))


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    main()
