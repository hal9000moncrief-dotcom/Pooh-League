import os
import re
import sys
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
        raise SystemExit("Usage: python app/build_summary_to_date.py [PD7]")
    return int(m.group(1))

def parse_owner_totals_from_final_owners_html(path: str) -> dict[str, int]:
    """
    Expects Final_Owners_PDx.html table:
    Owner | Starter Pooh Total | Starters Count So Far
    """
    with open(path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    table = soup.find("table")
    if not table:
        return {}

    rows = table.find_all("tr")
    out = {}
    for tr in rows[1:]:
        tds = tr.find_all("td")
        if len(tds) < 2:
            continue
        owner = tds[0].get_text(strip=True)
        total = tds[1].get_text(strip=True)
        try:
            out[owner] = int(total)
        except:
            out[owner] = 0
    return out

def main():
    cap = cap_pd_number(sys.argv[1] if len(sys.argv) > 1 else None)

    files = [f for f in os.listdir(DOCS_DIR) if f.startswith("Final_Owners_PD") and f.endswith(".html")]
    items = []
    for f in files:
        n = pd_num_from_name(f)
        if n < 0:
            continue
        if cap is not None and n > cap:
            continue
        items.append((n, f))
    items.sort()

    # Build per-PD totals + cumulative
    per_pd = {}
    owners_all = set()

    for n, f in items:
        path = os.path.join(DOCS_DIR, f)
        totals = parse_owner_totals_from_final_owners_html(path)
        per_pd[n] = totals
        owners_all |= set(totals.keys())

    owners = sorted(owners_all)

    cumulative = {o: 0 for o in owners}

    # HTML output
    out_path = os.path.join(DOCS_DIR, "SummaryToDate.html")
    with open(out_path, "w", encoding="utf-8") as out:
        out.write("<!doctype html><html><head><meta charset='utf-8'>")
        title = f"Summary To Date (through PD{cap})" if cap is not None else "Summary To Date"
        out.write(f"<title>{title}</title>")
        out.write("<style>body{font-family:Arial}table{border-collapse:collapse;font-size:14px}"
                  "th,td{border:1px solid #ccc;padding:4px 6px}th{background:#eee}</style>")
        out.write("</head><body>")
        out.write(f"<h2>{title}</h2>")

        out.write("<table><thead><tr><th>PD</th>")
        for o in owners:
            out.write(f"<th>{o}</th>")
        out.write("</tr></thead><tbody>")

        for n, _f in items:
            out.write(f"<tr><td>PD{n}</td>")
            totals = per_pd.get(n, {})
            for o in owners:
                v = int(totals.get(o, 0))
                cumulative[o] += v
                out.write(f"<td>{v}</td>")
            out.write("</tr>")

        out.write("<tr><th>CUM</th>")
        for o in owners:
            out.write(f"<th>{cumulative[o]}</th>")
        out.write("</tr>")

        out.write("</tbody></table></body></html>")

    print(f"Wrote: {out_path}")

if __name__ == "__main__":
    main()
