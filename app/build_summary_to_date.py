import os
import re
import sys
from collections import defaultdict
from bs4 import BeautifulSoup

from openpyxl import load_workbook

DOCS_DIR = os.path.join(os.path.dirname(__file__), "..", "docs")
TEAM_NAMES_XLSX = os.path.join(DOCS_DIR, "Team_Names.xlsx")


def parse_cap_pd(argv) -> int | None:
    # optional: PD7
    if len(argv) < 2:
        return None
    s = argv[1].strip().upper()
    m = re.fullmatch(r"PD(\d+)", s)
    if not m:
        raise SystemExit("Usage: python app/build_summary_to_date.py [PD7]")
    return int(m.group(1))


def pd_num_from_filename(fn: str) -> int | None:
    m = re.search(r"Final_Owners_PD(\d+)\.html$", fn)
    return int(m.group(1)) if m else None


def read_owner_totals_from_final_owners_html(path: str) -> dict[str, int]:
    """
    Reads docs/Final_Owners_PDx.html table like:
      Owner | Starter Pooh Total | Starters Count So Far
    Returns {owner: starter_pooh_total}.
    """
    with open(path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    table = soup.find("table")
    if not table:
        return {}

    rows = table.find_all("tr")
    out: dict[str, int] = {}

    for tr in rows[1:]:
        tds = tr.find_all("td")
        if len(tds) < 2:
            continue
        owner = tds[0].get_text(strip=True)
        total_txt = tds[1].get_text(strip=True)
        try:
            out[owner] = int(total_txt)
        except:
            out[owner] = 0

    return out


# ----------------------------
# Team Names Map (Owner -> Team Name)
# ----------------------------
def load_team_name_map() -> dict[str, str]:
    """
    docs/Team_Names.xlsx:
      headers: Owner, Team Name
    """
    if not os.path.exists(TEAM_NAMES_XLSX):
        print(f"NOTE: Missing {TEAM_NAMES_XLSX}. Using names as-is.")
        return {}

    wb = load_workbook(TEAM_NAMES_XLSX, data_only=True)
    ws = wb.active

    headers = [("" if c.value is None else str(c.value).strip()) for c in ws[1]]
    headers_l = [h.lower() for h in headers]

    def col(name: str) -> int | None:
        n = name.lower()
        if n in headers_l:
            return headers_l.index(n) + 1
        return None

    c_owner = col("Owner")
    c_team = col("Team Name")
    if not c_owner or not c_team:
        raise SystemExit("ERROR: docs/Team_Names.xlsx must have headers: Owner, Team Name")

    m: dict[str, str] = {}
    for r in range(2, ws.max_row + 1):
        old = ws.cell(row=r, column=c_owner).value
        new = ws.cell(row=r, column=c_team).value
        old_s = "" if old is None else str(old).strip()
        new_s = "" if new is None else str(new).strip()
        if old_s and new_s:
            m[old_s] = new_s

    return m


def display_team(name: str, team_map: dict[str, str]) -> str:
    s = (name or "").strip()
    if not s:
        return s
    if s == "Undrafted":
        return s
    return team_map.get(s, s)


def main():
    cap_pd = parse_cap_pd(sys.argv)

    team_map = load_team_name_map()
    if team_map:
        print(f"Loaded Team_Names mapping entries: {len(team_map)}")
    else:
        print("No Team_Names mapping loaded (using names as-is).")

    # Find Final_Owners_PD*.html
    pd_files: list[tuple[int, str]] = []
    for fn in os.listdir(DOCS_DIR):
        n = pd_num_from_filename(fn)
        if n is None:
            continue
        if cap_pd is not None and n > cap_pd:
            continue
        pd_files.append((n, fn))

    pd_files.sort(key=lambda x: x[0])  # PD1..PDN

    if not pd_files:
        raise SystemExit("No Final_Owners_PD*.html files found in docs/")

    max_pd = pd_files[-1][0]

    # per_owner_per_pd[owner][pd] = points for that PD
    per_owner_per_pd: dict[str, dict[int, int]] = defaultdict(dict)
    owners_set = set()

    for pd, fn in pd_files:
        path = os.path.join(DOCS_DIR, fn)
        totals = read_owner_totals_from_final_owners_html(path)
        for owner, v in totals.items():
            owners_set.add(owner)
            per_owner_per_pd[owner][pd] = int(v)

    owners = sorted(list(owners_set))

    # Totals + avg
    pd_list = [pd for pd, _ in pd_files]
    completed_pd_count = len(pd_list)

    owner_total: dict[str, int] = {}
    owner_avg: dict[str, float] = {}

    for owner in owners:
        pd_scores = [per_owner_per_pd[owner].get(pd, 0) for pd in pd_list]
        total = sum(pd_scores)
        owner_total[owner] = total
        owner_avg[owner] = (total / completed_pd_count) if completed_pd_count > 0 else 0.0

    # Sort by Total Pooh descending (keep stable), tie-break by display name
    owners_sorted = sorted(
        owners,
        key=lambda o: (-owner_total.get(o, 0), display_team(o, team_map))
    )

    # Reference totals for Out Of 1st/2nd/3rd
    top1 = owner_total.get(owners_sorted[0], 0) if len(owners_sorted) >= 1 else 0
    top2 = owner_total.get(owners_sorted[1], top1) if len(owners_sorted) >= 2 else top1
    top3 = owner_total.get(owners_sorted[2], top2) if len(owners_sorted) >= 3 else top2

    # Write SummaryToDate.html with your headers
    out_path = os.path.join(DOCS_DIR, "SummaryToDate.html")

    with open(out_path, "w", encoding="utf-8") as out:
        out.write("<!doctype html><html><head><meta charset='utf-8'>")
        out.write("<title>Sorted League Results</title>")
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
        out.write("<h2 style='text-align:center'>Sorted League Results</h2>")

        out.write("<table><thead><tr>")
        out.write("<th>Team Name</th>")
        out.write("<th>Total Pooh</th>")
        out.write("<th>Out Of 1st</th>")
        out.write("<th>Out Of 2nd</th>")
        out.write("<th>Out Of 3rd</th>")

        # PD columns: 1..max_pd (not hardcoded 19)
        for pd in range(1, max_pd + 1):
            out.write(f"<th>{pd}</th>")

        out.write("<th>Avg Pooh Per Completed PD</th>")

        # Keep this column but leave it blank for every row (per your instruction)
        out.write("<th>Sum of Avgs, Top 5 Eligible</th>")

        # DO NOT include "Remaining Current PD"
        out.write("</tr></thead><tbody>")

        for owner in owners_sorted:
            total = owner_total.get(owner, 0)

            out1 = max(0, top1 - total)
            out2 = max(0, top2 - total)
            out3 = max(0, top3 - total)

            owner_disp = display_team(owner, team_map)

            out.write("<tr>")
            out.write(f"<td>{owner_disp}</td>")
            out.write(f"<td class='num'>{total}</td>")
            out.write(f"<td class='num'>{out1}</td>")
            out.write(f"<td class='num'>{out2}</td>")
            out.write(f"<td class='num'>{out3}</td>")

            for pd in range(1, max_pd + 1):
                out.write(f"<td class='num'>{per_owner_per_pd[owner].get(pd, 0)}</td>")

            out.write(f"<td class='num'>{owner_avg.get(owner, 0.0):.2f}</td>")

            # Blank column on purpose
            out.write("<td class='num'></td>")

            out.write("</tr>")

        out.write("</tbody></table></body></html>")

    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
