import os
import re
import sys
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

from bs4 import BeautifulSoup
from openpyxl import load_workbook

DOCS_DIR = os.path.join(os.path.dirname(__file__), "..", "docs")
APP_DIR  = os.path.join(os.path.dirname(__file__), "..", "app")

ROSTERS_XLSX = os.path.join(APP_DIR, "rosters.xlsx")

# ----------------------------
# Helpers
# ----------------------------
def parse_cap_pd(argv) -> Optional[int]:
    # optional arg: PD7
    if len(argv) < 2:
        return None
    s = argv[1].strip().upper()
    m = re.fullmatch(r"PD(\d+)", s)
    if not m:
        raise SystemExit("Usage: python app/build_player_pooh_summary.py [PD7]")
    return int(m.group(1))

def pd_num_from_filename(fn: str) -> Optional[int]:
    m = re.search(r"Final_Players_PD(\d+)\.html$", fn)
    return int(m.group(1)) if m else None

def norm_name(name: str) -> str:
    s = (name or "").lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\b(jr|sr|ii|iii|iv)\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def safe_int(x) -> int:
    try:
        return int(str(x).strip())
    except:
        return 0

def safe_float(x) -> float:
    try:
        return float(str(x).strip())
    except:
        return 0.0

def html_read_table(path: str) -> Tuple[List[str], List[List[str]]]:
    with open(path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    table = soup.find("table")
    if not table:
        return [], []

    headers = [th.get_text(strip=True) for th in table.find_all("th")]
    headers_l = [h.lower() for h in headers]

    rows = []
    for tr in table.find_all("tr")[1:]:
        tds = tr.find_all("td")
        if not tds:
            continue
        rows.append([td.get_text(strip=True) for td in tds])

    return headers_l, rows

def idx(headers_l: List[str], *cands: str) -> Optional[int]:
    for c in cands:
        c = c.lower()
        if c in headers_l:
            return headers_l.index(c)
    return None

# ----------------------------
# Load rosters.xlsx (bio fields)
# ----------------------------
def load_rosters() -> Dict[str, dict]:
    """
    Your rosters.xlsx headers (per screenshot):
      Name, Order, Cost, Owner, Team, Height, Weight, Class, Position
    We'll treat Owner as "Team Name".
    """
    if not os.path.exists(ROSTERS_XLSX):
        raise SystemExit(f"ERROR: Missing {ROSTERS_XLSX}")

    wb = load_workbook(ROSTERS_XLSX, data_only=True)
    ws = wb.active

    headers = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
    headers_l = [h.lower() for h in headers]

    def col(*cands):
        for c in cands:
            if c.lower() in headers_l:
                return headers_l.index(c.lower()) + 1
        return None

    c_name = col("name", "player")
    if not c_name:
        raise SystemExit("ERROR: rosters.xlsx must have a 'Name' column.")

    c_cost   = col("cost")
    c_owner  = col("owner", "team name")
    c_team   = col("team")
    c_height = col("height")
    c_weight = col("weight")
    c_class  = col("class")
    c_pos    = col("position")

    out: Dict[str, dict] = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=c_name).value
        name = "" if name is None else str(name).strip()
        if not name:
            continue
        key = norm_name(name)

        out[key] = {
            "Name": name,
            "Cost": "" if c_cost is None else ("" if ws.cell(row=r, column=c_cost).value is None else str(ws.cell(row=r, column=c_cost).value).strip()),
            "Team Name": "" if c_owner is None else ("" if ws.cell(row=r, column=c_owner).value is None else str(ws.cell(row=r, column=c_owner).value).strip()),
            "Team": "" if c_team is None else ("" if ws.cell(row=r, column=c_team).value is None else str(ws.cell(row=r, column=c_team).value).strip()),
            "Height": "" if c_height is None else ("" if ws.cell(row=r, column=c_height).value is None else str(ws.cell(row=r, column=c_height).value).strip()),
            "Weight": "" if c_weight is None else ("" if ws.cell(row=r, column=c_weight).value is None else str(ws.cell(row=r, column=c_weight).value).strip()),
            "Class": "" if c_class is None else ("" if ws.cell(row=r, column=c_class).value is None else str(ws.cell(row=r, column=c_class).value).strip()),
            "Position": "" if c_pos is None else ("" if ws.cell(row=r, column=c_pos).value is None else str(ws.cell(row=r, column=c_pos).value).strip()),
        }

    return out

# ----------------------------
# Load Final_Players_PD*.html (pooh per PD + stat totals)
# ----------------------------
def load_final_player_data(cap_pd: Optional[int]):
    """
    Returns:
      max_pd
      pooh_by_player_pd[player_norm][pd] = pooh
      agg_stats[player_norm] = totals across included PDs:
         games, min, pts, reb, ast, stl, blk, to
      owner_by_player[player_norm] = owner (from Final files when available)
    """
    files = []
    for fn in os.listdir(DOCS_DIR):
        n = pd_num_from_filename(fn)
        if n is None:
            continue
        if cap_pd is not None and n > cap_pd:
            continue
        files.append((n, fn))
    files.sort(key=lambda x: x[0])

    if not files:
        raise SystemExit("ERROR: No docs/Final_Players_PD*.html found.")

    max_pd = files[-1][0]

    pooh_by_player_pd: Dict[str, Dict[int, int]] = defaultdict(dict)
    agg = defaultdict(lambda: {"games": 0, "min": 0.0, "pts": 0, "reb": 0, "ast": 0, "stl": 0, "blk": 0, "to": 0})
    owner_by_player: Dict[str, str] = {}

    for pd, fn in files:
        path = os.path.join(DOCS_DIR, fn)
        headers_l, rows = html_read_table(path)
        if not headers_l or not rows:
            continue

        i_owner  = idx(headers_l, "owner")
        i_player = idx(headers_l, "player")
        i_pooh   = idx(headers_l, "pooh")
        i_pts    = idx(headers_l, "pts")
        i_reb    = idx(headers_l, "reb")
        i_ast    = idx(headers_l, "ast")
        i_stl    = idx(headers_l, "stl")
        i_blk    = idx(headers_l, "blk")
        i_to     = idx(headers_l, "to")
        i_min    = idx(headers_l, "min")

        if i_player is None or i_pooh is None:
            continue

        for r in rows:
            if i_player >= len(r):
                continue
            pname = r[i_player]
            key = norm_name(pname)
            if not key:
                continue

            pooh = safe_int(r[i_pooh]) if i_pooh < len(r) else 0
            pooh_by_player_pd[key][pd] = pooh

            # If this row exists, count it as a game played for that PD
            # (Final_Players files should only have players who actually played)
            agg[key]["games"] += 1
            if i_min is not None and i_min < len(r):
                agg[key]["min"] += safe_float(r[i_min])
            if i_pts is not None and i_pts < len(r):
                agg[key]["pts"] += safe_int(r[i_pts])
            if i_reb is not None and i_reb < len(r):
                agg[key]["reb"] += safe_int(r[i_reb])
            if i_ast is not None and i_ast < len(r):
                agg[key]["ast"] += safe_int(r[i_ast])
            if i_stl is not None and i_stl < len(r):
                agg[key]["stl"] += safe_int(r[i_stl])
            if i_blk is not None and i_blk < len(r):
                agg[key]["blk"] += safe_int(r[i_blk])
            if i_to is not None and i_to < len(r):
                agg[key]["to"] += safe_int(r[i_to])

            if i_owner is not None and i_owner < len(r):
                ow = r[i_owner].strip()
                if ow:
                    owner_by_player[key] = ow

    return max_pd, pooh_by_player_pd, agg, owner_by_player

# ----------------------------
# Write Player_Pooh_Summary.html
# ----------------------------
def write_html(out_path: str, cols: List[str], rows: List[Dict[str, str]]):
    with open(out_path, "w", encoding="utf-8") as out:
        out.write("<!doctype html><html><head><meta charset='utf-8'>")
        out.write("<title>Player Pooh Summary</title>")
        out.write(
            "<style>"
            "body{font-family:Arial}"
            "table{border-collapse:collapse;font-size:14px}"
            "th,td{border:1px solid #ccc;padding:4px 6px}"
            "th{background:#eee}"
            "td.num{text-align:right}"
            "</style>"
        )
        out.write("</head><body>")
        out.write("<h2 style='text-align:center'>Player Pooh Summary</h2>")

        out.write("<table><thead><tr>")
        for c in cols:
            out.write(f"<th>{c}</th>")
        out.write("</tr></thead><tbody>")

        for r in rows:
            out.write("<tr>")
            for c in cols:
                v = r.get(c, "")
                is_num = (c in {"Cost","Min/G","Avg","Total","PPG","R/G","A/G","B/G","S/G","T/G"} or c.isdigit())
                cls = " class='num'" if is_num else ""
                out.write(f"<td{cls}>{v}</td>")
            out.write("</tr>")

        out.write("</tbody></table></body></html>")

    print(f"Wrote: {out_path}")

# ----------------------------
# Main
# ----------------------------
def main():
    cap_pd = parse_cap_pd(sys.argv)

    rosters = load_rosters()
    max_pd, pooh_by_player_pd, agg, owner_by_player = load_final_player_data(cap_pd)

    # Columns EXACTLY as you requested
    fixed_cols = ["Team Name","Cost","Name","Team","Height","Weight","Class","Position","Min/G","Avg","Total"]
    pd_cols = [str(i) for i in range(1, max_pd + 1)]
    tail_cols = ["PPG","R/G","A/G","B/G","S/G","T/G"]
    cols = fixed_cols + pd_cols + tail_cols

    rows_out: List[Dict[str, str]] = []

    # Build rows from roster list (so you keep every drafted player even if they never played)
    for key, info in rosters.items():
        # Pooh per PD
        pd_vals = [pooh_by_player_pd.get(key, {}).get(pd, 0) for pd in range(1, max_pd + 1)]
        total_pooh = sum(pd_vals)
        avg_pooh = (total_pooh / max_pd) if max_pd > 0 else 0.0

        g = agg.get(key, {"games": 0, "min": 0.0, "pts": 0, "reb": 0, "ast": 0, "stl": 0, "blk": 0, "to": 0})
        games = g["games"] if g else 0

        def per_game(n: float) -> float:
            return (n / games) if games > 0 else 0.0

        min_g = per_game(g["min"])
        ppg   = per_game(g["pts"])
        rpg   = per_game(g["reb"])
        apg   = per_game(g["ast"])
        bpg   = per_game(g["blk"])
        spg   = per_game(g["stl"])
        tpg   = per_game(g["to"])

        team_name = owner_by_player.get(key) or info.get("Team Name", "")

        row = {
            "Team Name": team_name,
            "Cost": info.get("Cost",""),
            "Name": info.get("Name",""),
            "Team": info.get("Team",""),
            "Height": info.get("Height",""),
            "Weight": info.get("Weight",""),
            "Class": info.get("Class",""),
            "Position": info.get("Position",""),
            "Min/G": f"{min_g:.1f}",
            "Avg": f"{avg_pooh:.2f}",
            "Total": str(total_pooh),
            "PPG": f"{ppg:.2f}",
            "R/G": f"{rpg:.2f}",
            "A/G": f"{apg:.2f}",
            "B/G": f"{bpg:.2f}",
            "S/G": f"{spg:.2f}",
            "T/G": f"{tpg:.2f}",
        }
        for pd in range(1, max_pd + 1):
            row[str(pd)] = str(pooh_by_player_pd.get(key, {}).get(pd, 0))

        rows_out.append(row)

    # Sort by Total Pooh desc, then Name
    rows_out.sort(key=lambda r: (-safe_int(r.get("Total","0")), r.get("Name","")))

    out_path = os.path.join(DOCS_DIR, "Player_Pooh_Summary.html")
    write_html(out_path, cols, rows_out)

if __name__ == "__main__":
    main()
