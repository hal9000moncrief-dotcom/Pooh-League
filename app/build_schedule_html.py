import os
import html
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles.colors import Color

REPO_ROOT = os.path.join(os.path.dirname(__file__), "..")
DOCS_DIR = os.path.join(REPO_ROOT, "docs")

XLSX_PATH = os.path.join(DOCS_DIR, "Schedule 2026.xlsx")
OUT_HTML = os.path.join(DOCS_DIR, "Schedule.html")


# ----------------------------
# Theme color support helpers
# ----------------------------
def _hex_to_rgb(hex6: str):
    hex6 = hex6.strip().lstrip("#")
    return int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)

def _rgb_to_hex(r: int, g: int, b: int):
    return f"#{r:02X}{g:02X}{b:02X}"

def _apply_tint_to_rgb(r: int, g: int, b: int, tint: float):
    """
    Excel tint:
      - tint < 0 => darken
      - tint > 0 => lighten
    This approximates the Excel behavior.
    """
    tint = float(tint)
    def adj(c):
        if tint < 0:
            return int(round(c * (1.0 + tint)))
        else:
            return int(round(c + (255 - c) * tint))
    r2, g2, b2 = adj(r), adj(g), adj(b)
    r2 = max(0, min(255, r2))
    g2 = max(0, min(255, g2))
    b2 = max(0, min(255, b2))
    return r2, g2, b2

def _get_theme_palette_hex(wb) -> list:
    """
    Returns a list of up to 12 theme colors as hex strings 'RRGGBB'.
    If we can't parse the theme, we return a reasonable Office default palette.
    """
    # Office default (approx), indices 0..11
    fallback = [
        "FFFFFF",  # lt1
        "000000",  # dk1
        "EEECE1",  # lt2
        "1F497D",  # dk2
        "4F81BD",  # accent1
        "C0504D",  # accent2 (reddish)
        "9BBB59",  # accent3
        "8064A2",  # accent4
        "4BACC6",  # accent5
        "F79646",  # accent6
        "0000FF",  # hlink
        "800080",  # folHlink
    ]

    try:
        theme = wb.theme
        if theme is None:
            return fallback

        # openpyxl theme structure can vary; attempt to pull out clrScheme
        cs = theme.themeElements.clrScheme

        # Order should be: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink
        names = ["lt1","dk1","lt2","dk2","accent1","accent2","accent3","accent4","accent5","accent6","hlink","folHlink"]
        out = []

        for nm in names:
            cobj = getattr(cs, nm, None)
            if cobj is None:
                out.append(fallback[len(out)])
                continue

            # Theme colors often stored as srgbClr or sysClr
            val = None
            if getattr(cobj, "srgbClr", None) is not None:
                val = getattr(cobj.srgbClr, "val", None)
            if val is None and getattr(cobj, "sysClr", None) is not None:
                # sysClr has lastClr sometimes
                val = getattr(cobj.sysClr, "lastClr", None) or getattr(cobj.sysClr, "val", None)

            if not val:
                out.append(fallback[len(out)])
            else:
                # val already hex without '#'
                v = str(val).strip().lstrip("#")
                if len(v) == 6:
                    out.append(v.upper())
                else:
                    out.append(fallback[len(out)])

        return out
    except Exception:
        return fallback


def _css_color_from_openpyxl_color(c: Color, theme_palette_hex: list):
    """
    Convert openpyxl Color to CSS #RRGGBB.
    Supports:
      - rgb / ARGB
      - theme + tint
    """
    if c is None:
        return None

    # 1) Direct RGB (best case)
    rgb = getattr(c, "rgb", None)
    if rgb:
        rgb = str(rgb).strip()
        # ARGB -> drop alpha
        if len(rgb) == 8:
            rgb = rgb[2:]
        if len(rgb) == 6:
            return f"#{rgb.upper()}"
        return None

    # 2) Theme color
    theme_idx = getattr(c, "theme", None)
    if theme_idx is not None:
        try:
            idx = int(theme_idx)
        except Exception:
            return None

        if 0 <= idx < len(theme_palette_hex):
            base_hex = theme_palette_hex[idx]
        else:
            return None

        r, g, b = _hex_to_rgb(base_hex)
        tint = getattr(c, "tint", None)
        if tint is not None:
            r, g, b = _apply_tint_to_rgb(r, g, b, float(tint))
        return _rgb_to_hex(r, g, b)

    return None


# ----------------------------
# Excel -> HTML styling
# ----------------------------
def _cell_style_to_css(cell, theme_palette_hex):
    styles = []

    # Font
    f = cell.font
    if f is not None:
        if f.bold:
            styles.append("font-weight:700")
        if f.italic:
            styles.append("font-style:italic")
        if f.underline:
            styles.append("text-decoration:underline")
        if f.color is not None:
            col = _css_color_from_openpyxl_color(f.color, theme_palette_hex)
            if col:
                styles.append(f"color:{col}")

    # Fill (background)
    fill = cell.fill
    if fill is not None and getattr(fill, "patternType", None) == "solid":
        fg = getattr(fill, "fgColor", None)
        col = _css_color_from_openpyxl_color(fg, theme_palette_hex)
        if col:
            styles.append(f"background:{col}")

    # Alignment
    a = cell.alignment
    if a is not None:
        if a.horizontal:
            styles.append(f"text-align:{a.horizontal}")
        if a.vertical:
            styles.append(f"vertical-align:{a.vertical}")
        if a.wrap_text:
            styles.append("white-space:pre-wrap")

    return ";".join(styles)


def _escape_cell_value(v):
    if v is None:
        return "&nbsp;"
    s = str(v)
    s = html.escape(s)
    s = s.replace("\n", "<br>")
    if s.strip() == "":
        return "&nbsp;"
    return s


def main():
    if not os.path.isfile(XLSX_PATH):
        raise SystemExit(f"ERROR: Missing file: {XLSX_PATH}")

    wb = load_workbook(XLSX_PATH, data_only=True)
    ws = wb.active

    theme_palette_hex = _get_theme_palette_hex(wb)

    max_row = ws.max_row or 1
    max_col = ws.max_column or 1

    # merged-cell maps: top-left -> (rowspan, colspan), covered cells skipped
    merged_top_left = {}
    merged_covered = set()
    for m in ws.merged_cells.ranges:
        min_row = m.min_row
        min_col = m.min_col
        max_row_m = m.max_row
        max_col_m = m.max_col
        rs = max_row_m - min_row + 1
        cs = max_col_m - min_col + 1
        merged_top_left[(min_row, min_col)] = (rs, cs)
        for r in range(min_row, max_row_m + 1):
            for c in range(min_col, max_col_m + 1):
                if (r, c) != (min_row, min_col):
                    merged_covered.add((r, c))

    updated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write("<!doctype html><html><head><meta charset='utf-8'>")
        f.write("<title>Schedule</title>")
        f.write(
            "<style>"
            "html,body{margin:0;padding:0}"
            "body{font-family:Calibri,Arial;background:#ffffff}"
            ".wrap{max-width:98vw;margin:10px auto;border:3px solid #000;background:#FFFFCC;padding:10px;box-sizing:border-box}"
            "h1{margin:0 0 10px 0;text-align:center;background:#C0C0C0;border:1px solid #000;padding:10px}"
            "table{border-collapse:collapse;width:100%;table-layout:fixed;background:#ffffff}"
            "td,th{border:1px solid #000;padding:4px 6px;font-size:10pt;line-height:1.15;"
            "white-space:normal;word-break:break-word;overflow-wrap:anywhere}"
            "</style>"
        )
        f.write("</head><body><div class='wrap'>")
        f.write("<h1>Schedule</h1>")
        f.write(f"<div style='font-size:10pt;margin-bottom:10px;'><b>Last updated:</b> {html.escape(updated)}</div>")
        f.write("<table>")

        for r in range(1, max_row + 1):
            f.write("<tr>")
            for c in range(1, max_col + 1):
                if (r, c) in merged_covered:
                    continue

                cell = ws.cell(row=r, column=c)
                tag = "th" if r <= 2 else "td"  # 2 header rows

                attrs = []
                if (r, c) in merged_top_left:
                    rs, cs = merged_top_left[(r, c)]
                    if rs > 1:
                        attrs.append(f"rowspan='{rs}'")
                    if cs > 1:
                        attrs.append(f"colspan='{cs}'")

                css = _cell_style_to_css(cell, theme_palette_hex)
                if css:
                    attrs.append(f"style='{css}'")

                val = _escape_cell_value(cell.value)
                f.write(f"<{tag} {' '.join(attrs)}>{val}</{tag}>")

            f.write("</tr>")

        f.write("</table>")

        f.write("<div style='margin-top:10px;font-size:10pt;'>")
        f.write(f"<a href='{html.escape(os.path.basename(XLSX_PATH))}'>Download the Excel version</a>")
        f.write("</div>")

        f.write("</div></body></html>")

    print(f"Wrote: {OUT_HTML}")


if __name__ == "__main__":
    main()
