#!/usr/bin/env python3
"""Add Nintendo's latest official headcount to data/Nintendo.xlsx.

Nintendo updates its company profile every quarter, with its financial results
(early Feb, May, Aug, Nov). This script reads that page and, when it shows a
quarter the sheet doesn't have yet, adds an "Employees (MM/YYYY)" column just
before "Source" with the group and Nintendo Co., Ltd. figures (green = company info).

Usage:  python3 scripts/update_nintendo.py [--html saved_page.html]
Exit code 0 always; prints "updated" when the sheet changed, "unchanged" otherwise.
"""
import copy, pathlib, re, sys, urllib.request
import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "Nintendo.xlsx"
URL = "https://www.nintendo.co.jp/corporate/en/outline/index.html"
MONTHS = {m: i for i, m in enumerate(["January", "February", "March", "April", "May", "June", "July",
                                      "August", "September", "October", "November", "December"], 1)}
ROWS = {"group": "Nintendo Global", "ncl": "NCL (Nintendo Co. Ltd.)"}  # names in column C


def read_page():
    if "--html" in sys.argv:
        return pathlib.Path(sys.argv[sys.argv.index("--html") + 1]).read_text(encoding="utf-8")
    req = urllib.request.Request(URL, headers={"User-Agent": "SecondPartyBot/1.0 (+github.com)"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode("utf-8", "replace")


def parse(html):
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text)
    out = {}
    for key, pat in (("group", r"global consolidated basis"), ("ncl", r"Nintendo Co\., Ltd\. only")):
        m = re.search(r"([\d,]+)\s*\(\s*" + pat + r"\s+as of the end of (\w+) (\d{4})\s*\)", text)
        if not m:
            raise SystemExit(f"Could not find the {key} headcount on {URL}; the page layout may have changed.")
        out[key] = (int(m.group(1).replace(",", "")), MONTHS[m.group(2)], int(m.group(3)))
    if out["group"][1:] != out["ncl"][1:]:
        raise SystemExit("Group and Nintendo Co., Ltd. figures are for different dates; check the page by hand.")
    return out


def main():
    data = parse(read_page())
    _, month, year = data["group"]
    label = f"Employees ({month:02d}/{year})"
    wb = openpyxl.load_workbook(XLSX)
    ws = wb["Nintendo"]
    src = next(c for c in range(5, 60) if str(ws.cell(2, c).value or "").strip() == "Source")
    heads = {str(ws.cell(2, c).value).strip(): c for c in range(5, src)}
    rows = {k: next(r for r in range(3, 80) if ws.cell(r, 3).value == name) for k, name in ROWS.items()}

    if label in heads:
        col = heads[label]
        if all(ws.cell(rows[k], col).value == data[k][0] for k in rows):
            print("unchanged", label)
            return
    else:
        # Insert a new column just before "Source"; keep hyperlinks in the shifted columns
        links = {(c.row, c.column): c.hyperlink.target for row in ws.iter_rows() for c in row if c.hyperlink}
        width = ws.column_dimensions[openpyxl.utils.get_column_letter(src - 1)].width
        ws.insert_cols(src)
        for (r, c), t in links.items():
            if c >= src:
                ws.cell(r, c).hyperlink = None
                ws.cell(r, c + 1).hyperlink = t
        col = src
        ws.cell(2, col).value = label
        ws.cell(2, col)._style = copy.copy(ws.cell(2, col - 1)._style)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    for k, r in rows.items():
        c = ws.cell(r, col)
        c.value = data[k][0]
        c._style = copy.copy(ws.cell(r, col - 1)._style)  # same green "company info" fill as the previous figure
    wb.save(XLSX)
    print("updated", label, data["group"][0], data["ncl"][0])


if __name__ == "__main__":
    main()
