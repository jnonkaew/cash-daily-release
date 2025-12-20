import json
import urllib.request

VERSION_FILE = "version.json"
REMOTE_VERSION_URL = (
    "https://raw.githubusercontent.com/jnonkaew/cash-daily-release/main/version.json"
)

def get_local_version():
    try:
        with open(VERSION_FILE, "r", encoding="utf-8") as f:
            return json.load(f)["version"]
    except:
        return "0.0.0"

def get_remote_info():
    with urllib.request.urlopen(REMOTE_VERSION_URL, timeout=5) as r:
        return json.loads(r.read().decode())

def is_newer(local, remote):
    def to_tuple(v):
        return tuple(map(int, v.split(".")))
    return to_tuple(remote) > to_tuple(local)
