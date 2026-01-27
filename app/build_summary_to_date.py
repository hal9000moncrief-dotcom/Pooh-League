import os
import re
import html
from datetime import datetime
from bs4 import BeautifulSoup

DOCS_DIR = os.path.join(os.path.dirname(__file__), "..", "docs")
OUT_HTML = os.path.join(DOCS_DIR, "SummaryToDate.html")

# We will aggregate from these snapshot files created by Final runs
FINAL_OWNERS_RE = re.compile(r"^Final_Owners_(PD(\d+))\.html$", re.IGNORECASE)

def parse_int(s: str) -> int:
    try:
        return int(float(str(s).strip()))
    except:
        return 0

def extract_owner_rows(final_owners_html_path: str):
    """
    Reads one Final_Owners_PDx.html and returns list of tuples:
      (owner, starter_pooh_total, starters_count_so_far)
    This matches the table produced by python_today_pooh.py in write_html_tables().
    """
    with open(final_owners_html_path, "r", encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f.read(), "lxml")

    table = soup.find("table")
    if not table:
        return []

    rows = table.find_all("tr")
    out = []

    # First row is header in your generated owners page:
    # Owner | Starter Pooh Total | Starters Count So Far
    for tr in rows[1:]:
        tds = tr.find_all(["td", "th"])
        if len(tds) < 3:
            continue
        owner = tds[0].get_text(strip=True)
        pooh = parse_int(tds[1].get_text(strip=True))
        cnt  = parse_int(tds[2].get_text(strip=True))
        if owner:
            out.append((owner, pooh, cnt))

    return out

def main():
    os.makedirs(DOCS_DIR, exist_ok=True)

    # Find all Final_Owners_PDx.html files and sort by PD number
    finals = []
    for fn in os.listdir(DOCS_DIR):
        m = FINAL_OWNERS_RE.match(fn)
        if m:
            pd_label = m.group(1).upper()      # e.g. PD7
            pd_num = int(m.group(2))           # 7
            finals.append((pd_num, pd_label, os.path.join(DOCS_DIR, fn)))

    finals.sort(key=lambda x: x[0])  # by PD number ascending

    # If no finals yet, still create a page (prevents 404 from index)
    if not finals:
        with open(OUT_HTML, "w", encoding="utf-8") as f:
            f.write("<!doctype html><html><head><meta charset='utf-8'>")
            f.write("<title>Summary of Results to Date</title>")
            f.write("<style>body{font-family:Calibri,Arial} .wrap{width:1100px;margin:20px auto;"
                    "border:3px solid #000;background:#FFFFCC;padding:10px}</style>")
            f.write("</head><body><div class='wrap'>")
            f.write("<h2>Summary of Results to Date</h2>")
            f.write("<p>No Final PD snapshot files found yet (expected docs/Final_Owners_PD#.html).</p>")
            f.write("</div></body></html>")
        print(f"Wrote: {OUT_HTML} (no finals found)")
        return

    # Determine max PD present (columns 1..max_pd)
    max_pd = finals[-1][0]

    # Aggregation structure:
    # owner -> {
    #   "per_pd": {pd_num: pooh_for_that_pd},
    #   "total": sum,
    #   "avg": total / number_of_pds_with_value
    # }
    owners = {}

    for pd_num, pd_label, path in finals:
        rows = extract_owner_rows(path)
        for owner, pooh, _cnt in rows:
            if owner not in owners:
                owners[owner] = {"per_pd": {}, "total": 0}
            owners[owner]["per_pd"][pd_num] = pooh

    # Compute totals/avgs
    for owner, d in owners.items():
        per_pd = d["per_pd"]
        d["total"] = sum(per_pd.get(i, 0) for i in range(1, max_pd + 1))
        completed = sum(1 for i in range(1, max_pd + 1) if i in per_pd)
        d["completed"] = completed
        d["avg"] = (d["total"] / completed) if completed else 0.0

    # Rank by total desc
    ranked = sorted(owners.items(), key=lambda kv: kv[1]["total"], reverse=True)

    # Helper: "Out of 1st/2nd/3rd" computed from totals
    totals_list = [d["total"] for _, d in ranked]
    first = totals_list[0] if len(totals_list) >= 1 else 0
    second = totals_list[1] if len(totals_list) >= 2 else None
    third = totals_list[2] if len(totals_list) >= 3 else None

    def esc(x):
        return html.escape("" if x is None else str(x))

    # Write HTML styled like your example sheet output (Calibri, gray header, borders)
    # You said: drop last column, and second-to-last can be blank for now.
    # In your example, last 3 columns are:
    #  - Avg Pooh Per Completed PD
    #  - Sum of Avgs, Top 5 Eligible   (2nd-to-last)  <-- leave blank
    #  - Remaining Current PD         (last)          <-- omit entirely
    updated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write("<!doctype html><html><head><meta charset='utf-8'>")
        f.write("<title>Sorted League Results</title>")
        f.write("<style>"
                "body{font-family:Calibri,Arial;background:#ffffff}"
                "table{border-collapse:collapse;font-size:11pt}"
                "th,td{border:1px solid #000;padding:4px 6px}"
                "th{background:#c0c0c0}"
                "caption{font-weight:bold;margin-bottom:8px}"
                ".wrap{width:1400px;margin:20px auto}"
                "</style>")
        f.write("</head><body><div class='wrap'>")

        f.write("<table>")
        f.write("<caption>Sorted League Results</caption>")
        f.write(f"<tr><td colspan='{5 + max_pd + 2}' "
                "style='border:none;padding:0 0 10px 0;'>"
                f"<b>Included PDs:</b> 1–{max_pd} &nbsp;&nbsp; "
                f"<b>Last updated:</b> {esc(updated)}"
                "</td></tr>")

        # Header
        f.write("<thead><tr>")
        f.write("<th>Team Name</th>")
        f.write("<th>Total Pooh</th>")
        f.write("<th>Out Of 1st</th>")
        f.write("<th>Out Of 2nd</th>")
        f.write("<th>Out Of 3rd</th>")
        for i in range(1, max_pd + 1):
            f.write(f"<th>{i}</th>")
        f.write("<th>Avg Pooh Per Completed PD</th>")
        f.write("<th>Sum of Avgs, Top 5 Eligible</th>")  # blank for now
        f.write("</tr></thead><tbody>")

        # Rows
        for idx, (owner, d) in enumerate(ranked):
            total = d["total"]
            out1 = "---" if idx == 0 else str(first - total)
            out2 = "---" if idx <= 1 or second is None else str(second - total)
            out3 = "---" if idx <= 2 or third is None else str(third - total)

            f.write("<tr>")
            f.write(f"<td>{esc(owner)}</td>")
            f.write(f"<td style='text-align:right'>{total}</td>")
            f.write(f"<td style='text-align:right'>{esc(out1)}</td>")
            f.write(f"<td style='text-align:right'>{esc(out2)}</td>")
            f.write(f"<td style='text-align:right'>{esc(out3)}</td>")

            for i in range(1, max_pd + 1):
                val = d["per_pd"].get(i, "")
                f.write(f"<td style='text-align:right'>{esc(val)}</td>")

            f.write(f"<td style='text-align:right'>{d['avg']:.1f}</td>")
            f.write("<td style='text-align:right'></td>")  # blank for now
            f.write("</tr>")

        f.write("</tbody></table></div></body></html>")

    print(f"Wrote: {OUT_HTML}")

if __name__ == "__main__":
    main()
