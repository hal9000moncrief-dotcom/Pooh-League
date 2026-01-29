import os
import re
import sys
from collections import defaultdict
from bs4 import BeautifulSoup

DOCS_DIR = os.path.join(os.path.dirname(__file__), "..", "docs")

def pd_num_from_name(filename: str) -> int:
    m = re.search(r"_PD(\d+)\.html$", filename)
    return int(m.group(1)) if m else -1

def cap_pd_number(arg: str | None) -> int | None:
    if not arg:
        return None
    m = re.fullmatch(r"PD(\d+)", arg.strip().upper())
    if not m:
        raise SystemExit("Usage: python app/build_player_pooh_summary.py [PD7]")
    return int(m.group(1))

def parse_players_from_final_players_html(path: str):
    """
    Expects Final_Players_PDx.html columns include:
    owner, player, pooh
    """
    with open(path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    table = soup.find("table")
    if not table:
        return []

    # header index
    header = [th.get_text(strip=True).lower() for th in table.find_all("th")]
    def idx(col):
        try:
            return header.index(col)
        except:
            return None

    i_owner = idx("owner")
    i_player = idx("player")
    i_pooh = idx("pooh")

    if i_owner is None or i_player is None or i_pooh is None:
        return []

    out = []
    rows = table.find_all("tr")[1:]
    for tr in rows:
        tds = tr.find_all("td")
        if len(tds) <= max(i_owner, i_player, i_pooh):
            continue
        owner = tds[i_owner].get_text(strip=True)
        player = tds[i_player].get_text(strip=True)
        pooh_txt = tds[i_pooh].get_text(strip=True)
        try:
            pooh = int(pooh_txt)
        except:
            pooh = 0
        out.append((owner, player, pooh))
    return out

def main():
    cap = cap_pd_number(sys.argv[1] if len(sys.argv) > 1 else None)

    files = [f for f in os.listdir(DOCS_DIR) if f.startswith("Final_Players_PD") and f.endswith(".html")]
    items = []
    for f in files:
        n = pd_num_from_name(f)
        if n < 0:
            continue
        if cap is not None and n > cap:
            continue
        items.append((n, f))
    items.sort()

    total_by_player = defaultdict(int)
    owner_by_player = {}

    for n, f in items:
        path = os.path.join(DOCS_DIR, f)
        rows = parse_players_from_final_players_html(path)
        for owner, player, pooh in rows:
            # Keep last-seen owner; player names should be stable in your export
            owner_by_player[player] = owner
            total_by_player[player] += pooh

    ranked = sorted(total_by_player.items(), key=lambda x: x[1], reverse=True)

    out_path = os.path.join(DOCS_DIR, "Player_Pooh_Summary.html")
    with open(out_path, "w", encoding="utf-8") as out:
        out.write("<!doctype html><html><head><meta charset='utf-8'>")
        title = f"Player Pooh Summary (through PD{cap})" if cap is not None else "Player Pooh Summary"
        out.write(f"<title>{title}</title>")
        out.write("<style>body{font-family:Arial}table{border-collapse:collapse;font-size:14px}"
                  "th,td{border:1px solid #ccc;padding:4px 6px}th{background:#eee}</style>")
        out.write("</head><body>")
        out.write(f"<h2>{title}</h2>")
        out.write("<table><thead><tr><th>Rank</th><th>Player</th><th>Owner</th><th>Total Pooh</th></tr></thead><tbody>")

        for i, (player, total) in enumerate(ranked, 1):
            out.write(f"<tr><td>{i}</td><td>{player}</td><td>{owner_by_player.get(player,'')}</td><td>{total}</td></tr>")

        out.write("</tbody></table></body></html>")

    print(f"Wrote: {out_path}")

if __name__ == "__main__":
    main()
