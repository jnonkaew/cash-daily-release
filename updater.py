import urllib.request
import json
import time
import os
import sys

VERSION_URL = (
    "https://raw.githubusercontent.com/jnonkaew/cash-daily-release/main/version.json"
)

APP_EXE = "cash_daily.exe"

time.sleep(1)

with urllib.request.urlopen(VERSION_URL) as r:
    info = json.loads(r.read().decode())

exe_url = (
    f"https://raw.githubusercontent.com/jnonkaew/cash-daily-release/main/{info['exe']}"
)

tmp = "cash_daily_new.exe"
urllib.request.urlretrieve(exe_url, tmp)

time.sleep(1)

if os.path.exists(APP_EXE):
    os.remove(APP_EXE)

os.rename(tmp, APP_EXE)

with open("version.json", "w", encoding="utf-8") as f:
    json.dump({"version": info["version"]}, f)

os.startfile(APP_EXE)
sys.exit()
