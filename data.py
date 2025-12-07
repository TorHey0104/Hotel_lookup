import os
import json
import pandas as pd
from datetime import datetime

RECENT_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recent_configs.json")

def format_timestamp(path: str) -> str:
    try:
        mod_time = datetime.fromtimestamp(os.path.getmtime(path))
    except (FileNotFoundError, OSError):
        return "Unknown timestamp"
    return mod_time.strftime("%d.%m.%Y %H:%M")


def remember_config(path: str):
    try:
        if not path:
            return
        path = os.path.abspath(path)
        recent = []
        if os.path.isfile(RECENT_CONFIG_PATH):
            with open(RECENT_CONFIG_PATH, "r", encoding="utf-8") as fh:
                recent = json.load(fh)
        if path in recent:
            recent.remove(path)
        recent.insert(0, path)
        recent = recent[:10]
        with open(RECENT_CONFIG_PATH, "w", encoding="utf-8") as fh:
            json.dump(recent, fh, indent=2)
    except Exception:
        pass


def load_recent_configs():
    try:
        if os.path.isfile(RECENT_CONFIG_PATH):
            with open(RECENT_CONFIG_PATH, "r", encoding="utf-8") as fh:
                return json.load(fh)
    except Exception:
        return []
    return []
