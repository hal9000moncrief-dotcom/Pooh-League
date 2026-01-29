# --- SNIP: imports unchanged ---
import sys
import time
import random
import re
import requests
import os
import html
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from zoneinfo import ZoneInfo

BASE = "https://site.api.espn.com/apis/site/v2/sports/basketball/mens-college-basketball"
DRAFT_XLSX = os.path.join(os.path.dirname(__file__), "ByCoach.xlsx")

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json,text/plain,*/*",
    "Referer": "https://www.espn.com/",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

LOCAL_TZ = ZoneInfo("America/Chicago")

# -------------------------------------------------
# STATUS FORMATTER (THIS IS THE FIX)
# -------------------------------------------------
def format_event_status(event: dict) -> str:
    """
    Returns:
      - '2:10 2nd'
      - 'Half'
      - 'Final'
      - scheduled start string
    """
    comps = event.get("competitions", [])
    comp = comps[0] if comps else {}
    status = comp.get("status", {})
    stype = status.get("type", {})

    state = stype.get("state")  # pre | in | post

    if state == "post":
        return "Final"

    short = stype.get("shortDetail")
    if state == "in" and short:
        return short

    clock = status.get("displayClock")
    period = status.get("period")
    if state == "in" and clock and period:
        half = "1st" if period == 1 else "2nd"
        return f"{clock} {half}"

    return stype.get("detail") or "Scheduled"

# -------------------------------------------------
# SCOREBOARD
# -------------------------------------------------
def get_sec_events(date_yyyymmdd: str) -> List[dict]:
    url = f"{BASE}/scoreboard?dates={date_yyyymmdd}&groups=23&limit=500"
    return SESSION.get(url, timeout=30).json().get("events", [])

def extract_event_header(e: dict) -> dict:
    comps = e.get("competitions") or []
    comp = comps[0] if comps else {}

    competitors = comp.get("competitors") or []
    ha = {}
    for c in competitors:
        team = c.get("team", {})
        ha[c.get("homeAway")] = {
            "abbr": team.get("abbreviation", ""),
            "score": int(c.get("score") or 0),
        }

    return {
        "status": format_event_status(e),
        "home": ha.get("home", {}),
        "away": ha.get("away", {}),
    }

# -------------------------------------------------
# EVERYTHING BELOW THIS LINE IS UNCHANGED
# -------------------------------------------------
# (boxscore parsing, pooh math, output, etc.)
