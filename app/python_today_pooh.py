import sys
import time
import random
import re
import requests
import os
import html
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional

from bs4 import BeautifulSoup  # safe to keep (used elsewhere in your project)
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment

from zoneinfo import ZoneInfo


BASE = "https://site.api.espn.com/apis/site/v2/sports/basketball/mens-college-basketball"

DRAFT_XLSX = os.path.join(os.path.dirname(__file__), "ByCoach.xlsx")  # must be in same folder

# Team name mapping lives in docs/
TEAM_NAMES_XLSX = os.path.join(os.path.dirname(__file__), "..", "docs", "Team_Names.xlsx")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0 Safari/537.36",
    "Accept": "application/json,text/plain,*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.espn.com/",
    "Connection": "keep-alive",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

BASE_DELAY = 0.25
JITTER = 0.25
MAX_RETRIES = 6
TIMEOUT = 30

LOCAL_TZ = ZoneInfo("America/Chicago")
UTC_TZ = ZoneInfo("UTC")


# ----------------------------
# UTIL
# ----------------------------
def polite_sleep():
    time.sleep(BASE_DELAY + random.random() * JITTER)

def get_json(url: str) -> dict:
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = SESSION.get(url, timeout=TIMEOUT)
            r.raise_for_status()
            polite_sleep()
            return r.json()
        except Exception as e:
            last_err = e
            time.sleep((0.7 ** attempt) + random.random() * 0.7)
    raise RuntimeError(f"Failed after retries: {url}\nLast error: {last_err}")

def safe_int(v) -> int:
    try:
        return int(str(v).strip())
    except:
        return 0

def parse_made_attempt(s: str) -> Tuple[int, int]:
    try:
        a, b = str(s).split("-")
        return int(a), int(b)
    except:
        return 0, 0

def to_minutes(v) -> float:
    if v is None:
        return 0.0
    s = str(v).strip()
    if not s or s == "--":
        return 0.0
    if ":" in s:
        try:
            mm, ss = s.split(":")
            return int(mm) + int(ss) / 60.0
        except:
            return 0.0
    try:
        return float(s)
    except:
        return 0.0

def norm_name(name: str) -> str:
    s = (name or "").lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\b(jr|sr|ii|iii|iv)\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def compute_line(values: List[str], labels: List[str]) -> Optional[dict]:
    """
    Parse a single ESPN player stat row and return all fields needed for:
      - Pooh points
      - FG% / 3PT% / FT%
      - PF totals
    Pooh = PTS + REB + AST + STL + BLK - missedFG - missedFT - TO
    """
    if not labels or not values:
        return None

    def idx(*names: str) -> Optional[int]:
        for nm in names:
            try:
                return labels.index(nm)
            except ValueError:
                continue
        return None

    i_min = idx("MIN")
    i_fg  = idx("FG")
    i_3pt = idx("3PT", "3P")
    i_ft  = idx("FT")
    i_reb = idx("REB")
    i_ast = idx("AST")
    i_stl = idx("STL")
    i_blk = idx("BLK")
    i_to  = idx("TO")
    i_pf  = idx("PF")
    i_pts = idx("PTS")

    required = [i_min, i_reb, i_ast, i_stl, i_blk, i_to]
    if any(x is None for x in required):
        return None

    max_i = max([x for x in required if x is not None] + [i_fg or 0, i_ft or 0, i_3pt or 0, i_pf or 0, i_pts or 0])
    if len(values) <= max_i:
        return None

    mins = to_minutes(values[i_min]) if i_min is not None else 0.0

    fgm, fga = parse_made_attempt(values[i_fg]) if i_fg is not None else (0, 0)
    tpm, tpa = parse_made_attempt(values[i_3pt]) if i_3pt is not None else (0, 0)
    ftm, fta = parse_made_attempt(values[i_ft]) if i_ft is not None else (0, 0)

    pts = safe_int(values[i_pts]) if i_pts is not None else 0
    reb = safe_int(values[i_reb]) if i_reb is not None else 0
    ast = safe_int(values[i_ast]) if i_ast is not None else 0
    stl = safe_int(values[i_stl]) if i_stl is not None else 0
    blk = safe_int(values[i_blk]) if i_blk is not None else 0
    tov = safe_int(values[i_to])  if i_to  is not None else 0
    pf  = safe_int(values[i_pf])  if i_pf  is not None else 0

    # Skip truly blank/DNP rows
    if mins == 0 and pts == 0 and reb == 0 and ast == 0 and stl == 0 and blk == 0 and tov == 0 and pf == 0 and fga == 0 and fta == 0 and tpa == 0:
        return None

    missed_fg = max(0, fga - fgm)
    missed_ft = max(0, fta - ftm)

    pooh = (pts + reb + ast + stl + blk) - (missed_fg + missed_ft + tov)

    return {
        "MIN": mins,
        "PTS": pts, "REB": reb, "AST": ast, "STL": stl, "BLK": blk, "TO": tov, "PF": pf,
        "FGM": fgm, "FGA": fga,
        "3PM": tpm, "3PA": tpa,
        "FTM": ftm, "FTA": fta,
        "POOH": pooh,
    }

def parse_yyyymmdd(s: str) -> datetime:
    s = (s or "").strip()
    if not re.fullmatch(r"\d{8}", s):
        raise ValueError("Date must be YYYYMMDD (8 digits).")
    dt = datetime.strptime(s, "%Y%m%d")
    return dt.replace(tzinfo=LOCAL_TZ)

def fmt_yyyymmdd(dt: datetime) -> str:
    return dt.strftime("%Y%m%d")

def fmt_yyyy_mm_dd(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%d")

def event_local_yyyymmdd(e: dict) -> str:
    """
    ESPN event['date'] is ISO 8601, usually UTC (often ends with 'Z').
    Convert to America/Chicago and return YYYYMMDD.
    """
    iso = (e.get("date") or "").strip()
    if not iso:
        return ""
    if iso.endswith("Z"):
        iso = iso[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(iso)
    except Exception:
        return ""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC_TZ)
    return dt.astimezone(LOCAL_TZ).strftime("%Y%m%d")


# ----------------------------
# TEAM NAME MAP (Owner -> Team Name)
# ----------------------------
def load_team_name_map(xlsx_path: str) -> Dict[str, str]:
    """
    docs/Team_Names.xlsx with headers:
      Owner (old name), Team Name (new display name)
    """
    if not os.path.exists(xlsx_path):
        print(f"NOTE: Team name map not found at {xlsx_path}. Using Owner names as-is.")
        return {}

    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    headers = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
    headers_l = [h.lower() for h in headers]

    def col(name: str) -> Optional[int]:
        n = name.lower()
        if n in headers_l:
            return headers_l.index(n) + 1
        return None

    c_owner = col("Owner")
    c_team  = col("Team Name")
    if not c_owner or not c_team:
        raise SystemExit("ERROR: Team_Names.xlsx must have headers: Owner, Team Name")

    m: Dict[str, str] = {}
    for r in range(2, ws.max_row + 1):
        old = ws.cell(row=r, column=c_owner).value
        new = ws.cell(row=r, column=c_team).value
        old_s = "" if old is None else str(old).strip()
        new_s = "" if new is None else str(new).strip()
        if old_s and new_s:
            m[old_s] = new_s

    return m

def display_owner(owner: str, team_map: Dict[str, str]) -> str:
    if not owner:
        return owner
    if owner == "Undrafted":
        return owner
    return team_map.get(owner, owner)


# ----------------------------
# DRAFT BOARD
# ----------------------------
def load_draft_board(xlsx_path: str, team_map: Dict[str, str]) -> Tuple[Dict[str, dict], List[str]]:
    wb = load_workbook(xlsx_path)
    ws = wb.active

    headers = {}
    for cell in ws[1]:
        h = str(cell.value).strip() if cell.value is not None else ""
        headers[h.lower()] = cell.column

    if "name" not in headers or "owner" not in headers:
        raise RuntimeError("ByCoach.xlsx must have columns: Name, Owner (and optional Started).")

    col_name = headers["name"]
    col_owner = headers["owner"]
    col_started = headers.get("started", None)

    draft_map: Dict[str, dict] = {}
    owner_order: List[str] = []

    for r in range(2, ws.max_row + 1):
        nm = ws.cell(row=r, column=col_name).value
        ow = ws.cell(row=r, column=col_owner).value
        st = ws.cell(row=r, column=col_started).value if col_started else None

        name = str(nm).strip() if nm else ""
        owner_raw = str(ow).strip() if ow else ""
        started_raw = str(st).strip().lower() if st is not None else ""
        started = started_raw in ("yes", "y", "true", "1")

        if not name:
            continue
        key = norm_name(name)

        if not owner_raw:
            owner_raw = "Undrafted"

        owner_disp = display_owner(owner_raw, team_map)

        if key not in draft_map:
            draft_map[key] = {"owner": owner_disp, "started": started, "raw_name": name}

        if owner_disp and owner_disp != "Undrafted" and owner_disp not in owner_order:
            owner_order.append(owner_disp)

    return draft_map, owner_order


# ----------------------------
# SCOREBOARD (SEC DIRECT)
# ----------------------------
def get_sec_events(date_yyyymmdd: str) -> List[dict]:
    url = f"{BASE}/scoreboard?dates={date_yyyymmdd}&groups=23&limit=500"
    data = get_json(url)
    events = data.get("events", []) or []

    # STRICT FILTER: only keep events whose actual local date matches the requested date.
    filtered = [e for e in events if event_local_yyyymmdd(e) == date_yyyymmdd]

    if len(filtered) != len(events):
        print(f"NOTE: Filtered scoreboard events for {date_yyyymmdd}: kept {len(filtered)} of {len(events)}")
        offenders = [e for e in events if event_local_yyyymmdd(e) != date_yyyymmdd]
        for e in offenders[:5]:
            print(f"  Dropped event {e.get('id')} local_date={event_local_yyyymmdd(e)} raw={e.get('date')}")

    return filtered

def extract_event_header(e: dict) -> dict:
    comps = e.get("competitions") or []
    comp = comps[0] if comps else {}

    status_obj = comp.get("status", {}) or {}
    status_type = status_obj.get("type", {}) or {}
    detail = status_type.get("detail") or status_type.get("description") or status_type.get("name") or "Unknown"

    competitors = comp.get("competitors") or []
    ha = {}
    for c in competitors:
        ha_key = c.get("homeAway") or ""
        team = c.get("team", {}) or {}
        ha[ha_key] = {
            "id": str(team.get("id") or ""),
            "abbr": team.get("abbreviation") or "",
            "name": team.get("displayName") or team.get("shortDisplayName") or "",
            "score": safe_int(c.get("score")),
        }

    return {"status": detail, "home": ha.get("home", {}), "away": ha.get("away", {})}


# ----------------------------
# BOXSCORE PARSE
# ----------------------------
def iter_athlete_rows(stat_group: dict) -> List[dict]:
    rows = []
    for key in ("athletes", "bench", "reserves"):
        v = stat_group.get(key)
        if isinstance(v, list):
            rows.extend(v)
    if not rows and isinstance(stat_group.get("athletes"), list):
        rows = stat_group["athletes"]
    return rows

def get_boxscore_players(event_id: str) -> List[dict]:
    url = f"{BASE}/summary?event={event_id}"
    data = get_json(url)

    box = data.get("boxscore") or {}
    players_sections = box.get("players") or []
    if not players_sections:
        return []

    out = []
    for ps in players_sections:
        team = ps.get("team", {}) or {}
        tabbr = team.get("abbreviation") or ""

        seen = set()
        for stat_group in ps.get("statistics") or []:
            labels = stat_group.get("labels") or []
            if not labels:
                continue

            for ath in iter_athlete_rows(stat_group):
                athlete = ath.get("athlete", {}) or {}
                aid = str(athlete.get("id") or "")
                pname = athlete.get("displayName") or athlete.get("shortName") or athlete.get("fullName") or "Unknown"
                values = ath.get("stats") or []

                if aid and aid in seen:
                    continue

                line = compute_line(values, labels)
                if not line:
                    continue

                if aid:
                    seen.add(aid)

                out.append({
                    "team": tabbr,
                    "player": pname,
                    "pooh": line["POOH"],
                    "pts": line["PTS"],
                    "reb": line["REB"],
                    "ast": line["AST"],
                    "stl": line["STL"],
                    "blk": line["BLK"],
                    "to":  line["TO"],
                    "pf":  line["PF"],
                    "min": line["MIN"],
                    # NEW: raw makes/attempts so Final snapshots can drive stat pages w/o ESPN
                    "fgm": line["FGM"], "fga": line["FGA"],
                    "3pm": line["3PM"], "3pa": line["3PA"],
                    "ftm": line["FTM"], "fta": line["FTA"],
                })

    return out


# ----------------------------
# OUTPUT HELPERS
# ----------------------------
def autosize_columns(ws):
    for col in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col)
        for row in range(1, ws.max_row + 1):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

def write_xlsx(players_rows: List[dict], owner_totals_rows: List[dict], out_path: str):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Players"

    headers1 = [
        "date","owner","started_today","player","team","game","status",
        "pooh","pts","reb","ast","stl","blk","to","pf","min",
        "fgm","fga","3pm","3pa","ftm","fta"
    ]
    ws1.append(headers1)
    for c in range(1, len(headers1) + 1):
        ws1.cell(row=1, column=c).font = Font(bold=True)
        ws1.cell(row=1, column=c).alignment = Alignment(horizontal="center")

    for r in players_rows:
        ws1.append([r.get(h, "") for h in headers1])

    ws2 = wb.create_sheet("OwnerTotals")
    headers2 = ["owner", "starter_pooh_total", "starters_count_so_far"]
    ws2.append(headers2)
    for c in range(1, len(headers2) + 1):
        ws2.cell(row=1, column=c).font = Font(bold=True)
        ws2.cell(row=1, column=c).alignment = Alignment(horizontal="center")

    for r in owner_totals_rows:
        ws2.append([r["owner"], r["starter_pooh_total"], r["starters_count_so_far"]])

    autosize_columns(ws1)
    autosize_columns(ws2)

    wb.save(out_path)

def write_html_tables(players_rows, owner_totals_rows, out_players_html, out_owners_html, title_str):
    def esc(x):
        return html.escape("" if x is None else str(x))

    # NEW columns added at end
    players_cols = [
        "date","owner","started_today","player","team","game","status",
        "pooh","pts","reb","ast","stl","blk","to","pf","min",
        "fgm","fga","3pm","3pa","ftm","fta"
    ]

    # Players page
    with open(out_players_html, "w", encoding="utf-8") as f:
        f.write("<!doctype html><html><head><meta charset='utf-8'>")
        f.write(f"<title>SEC Pooh Points — {esc(title_str)}</title>")
        f.write("<style>body{font-family:Arial}table{border-collapse:collapse;font-size:14px}"
                "th,td{border:1px solid #ccc;padding:4px 6px}th{background:#eee}"
                ".start{font-weight:bold}</style>")
        f.write("</head><body>")
        f.write(f"<h2>SEC Pooh Points — {esc(title_str)}</h2>")
        f.write("<table><thead><tr>")
        for c in players_cols:
            f.write(f"<th>{esc(c)}</th>")
        f.write("</tr></thead><tbody>")
        for r in players_rows:
            cls = " class='start'" if r.get("started_today") == "Yes" else ""
            f.write("<tr>")
            for c in players_cols:
                f.write(f"<td{cls}>{esc(r.get(c,''))}</td>")
            f.write("</tr>")
        f.write("</tbody></table></body></html>")

    # Owners page
    with open(out_owners_html, "w", encoding="utf-8") as f:
        f.write("<!doctype html><html><head><meta charset='utf-8'>")
        f.write(f"<title>Owner Starters Total — {esc(title_str)}</title>")
        f.write("<style>body{font-family:Arial}table{border-collapse:collapse;font-size:14px}"
                "th,td{border:1px solid #ccc;padding:4px 6px}th{background:#eee}</style>")
        f.write("</head><body>")
        f.write(f"<h2>Owner Starters Total — {esc(title_str)}</h2>")
        f.write("<table><thead><tr><th>Owner</th><th>Starter Pooh Total</th><th>Starters Count So Far</th></tr></thead><tbody>")
        for r in owner_totals_rows:
            f.write(f"<tr><td>{esc(r['owner'])}</td><td>{esc(r['starter_pooh_total'])}</td><td>{esc(r['starters_count_so_far'])}</td></tr>")
        f.write("</tbody></table></body></html>")


# ----------------------------
# MAIN
# ----------------------------
def main():
    # FINAL mode is set by workflow env RUN_MODE=final
    run_mode = (os.environ.get("RUN_MODE") or "normal").strip().lower()
    is_final = (run_mode == "final")

    # Primary date: from arg or "today" in America/Chicago
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        primary_dt = parse_yyyymmdd(sys.argv[1].strip())
    else:
        primary_dt = datetime.now(LOCAL_TZ)

    prev_dt = primary_dt - timedelta(days=1)

    primary_yyyymmdd = fmt_yyyymmdd(primary_dt)
    prev_yyyymmdd = fmt_yyyymmdd(prev_dt)

    primary_label = fmt_yyyy_mm_dd(primary_dt)
    prev_label = fmt_yyyy_mm_dd(prev_dt)

    # Title: always show range; add FINAL marker when applicable
    title_str = f"{prev_label} + {primary_label}"
    if is_final:
        title_str += " (FINAL snapshot)"

    # Load Owner -> Team Name mapping (forward-only)
    team_map = load_team_name_map(TEAM_NAMES_XLSX)
    if team_map:
        print(f"Loaded Team_Names mapping entries: {len(team_map)}")
    else:
        print("No Team_Names mapping loaded (using Owner names as-is).")

    print(f"RUN_MODE={run_mode}  is_final={is_final}")
    print(f"Local now (America/Chicago): {datetime.now(LOCAL_TZ).isoformat()}")
    print(f"Querying SEC scoreboard dates: {prev_yyyymmdd} and {primary_yyyymmdd}")

    draft_map, owner_order = load_draft_board(DRAFT_XLSX, team_map)

    output_dir = os.path.join(os.path.dirname(__file__), "..", "docs")
    os.makedirs(output_dir, exist_ok=True)

    out_xlsx = os.path.join(output_dir, f"Today_PoohPoints_SEC_ByOwner_{primary_label}.xlsx")
    out_players_html = os.path.join(output_dir, "today_players.html")
    out_owners_html  = os.path.join(output_dir, "today_owners.html")

    all_rows: List[dict] = []

    def process_day(day_label: str, date_yyyymmdd: str):
        nonlocal all_rows
        sec_events = get_sec_events(date_yyyymmdd)
        print(f"--- {day_label} ({date_yyyymmdd}) ---")
        print(f"SEC events after strict date filter: {len(sec_events)}")

        for e in sec_events:
            event_id = str(e.get("id") or "")
            hdr = extract_event_header(e)
            home = hdr["home"]
            away = hdr["away"]
            status_line = hdr["status"]

            game_label = f"{away.get('abbr','')}@{home.get('abbr','')}"
            print(f"{game_label} — {status_line} — (event {event_id})")

            players = get_boxscore_players(event_id)
            if not players:
                print("  (No boxscore player stats published yet — try again later.)\n")
                continue

            for p in players:
                key = norm_name(p["player"])
                info = draft_map.get(key)
                if info:
                    owner_disp = info["owner"]  # already mapped to Team Name
                    started_today = "Yes" if info["started"] else "No"
                else:
                    owner_disp = "Undrafted"
                    started_today = "No"

                all_rows.append({
                    "date": day_label,
                    "game": game_label,
                    "status": status_line,
                    "owner": owner_disp,
                    "started_today": started_today,
                    "team": p["team"],
                    "player": p["player"],
                    "pooh": p["pooh"],
                    "pts": p["pts"],
                    "reb": p["reb"],
                    "ast": p["ast"],
                    "stl": p["stl"],
                    "blk": p["blk"],
                    "to":  p["to"],
                    "pf":  p["pf"],
                    "min": p["min"],
                    "fgm": p["fgm"], "fga": p["fga"],
                    "3pm": p["3pm"], "3pa": p["3pa"],
                    "ftm": p["ftm"], "fta": p["fta"],
                })

            print(f"  Players captured: {len(players)}\n")

    # ALWAYS include previous day + primary day (this is your PD definition)
    process_day(prev_label, prev_yyyymmdd)
    process_day(primary_label, primary_yyyymmdd)

    # Sort players: Owner (draft order), then starters first, then Pooh desc
    owner_rank = {o: i for i, o in enumerate(owner_order)}

    def sort_key(r):
        o = r["owner"]
        oidx = owner_rank.get(o, 10_000 if o == "Undrafted" else 9_000)
        starter_rank = 0 if r["started_today"] == "Yes" else 1
        return (oidx, o, r.get("date",""), starter_rank, -int(r.get("pooh", 0)), r.get("player",""))

    all_rows.sort(key=sort_key)

    # OwnerTotals: EXCLUDE Undrafted
    totals: Dict[str, Dict[str, int]] = {}
    for r in all_rows:
        owner = r["owner"]
        if owner == "Undrafted":
            continue

        if owner not in totals:
            totals[owner] = {"starter_pooh_total": 0, "starters_count_so_far": 0}

        if r["started_today"] == "Yes":
            totals[owner]["starter_pooh_total"] += int(r["pooh"])
            totals[owner]["starters_count_so_far"] += 1

    owner_totals_rows = [{"owner": o, **vals} for o, vals in totals.items()]
    owner_totals_rows.sort(key=lambda x: x["starter_pooh_total"], reverse=True)

    write_xlsx(all_rows, owner_totals_rows, out_xlsx)
    write_html_tables(all_rows, owner_totals_rows, out_players_html, out_owners_html, title_str)

    print(f"Wrote: {out_players_html}")
    print(f"Wrote: {out_owners_html}")
    print(f"Wrote: {out_xlsx}")


if __name__ == "__main__":
    main()
