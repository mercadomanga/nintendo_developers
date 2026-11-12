#!/usr/bin/env python3
"""Build public/index.html from data/Nintendo.xlsx + data/extra.json + src/template.html.

Usage:  python3 scripts/build.py [--date YYYY-MM-DD] [--no-min] [--preview PATH]
Needs:  pip install openpyxl; npm install (terser, to minify the script)
"""
import os
import json, re, sys, datetime, pathlib, gzip, subprocess
from html import escape as html_escape
import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "Nintendo.xlsx"
EXTRA = ROOT / "data" / "extra.json"
TPL = ROOT / "src" / "template.html"
OUT = ROOT / "public" / "index.html"
# Public address of the site. On GitHub the workflow passes the Pages URL in SITE_URL;
# set it by hand if you build locally for another host or a custom domain.
SITE = (os.environ.get("SITE_URL") or "https://mercadomanga.github.io/nintendo_developers/").rstrip("/") + "/"
BSKY = "https://bsky.app/profile/nakuto.eurosky.social"

# Cell colours used in the sheet's legend: green = company info, yellow = third-party info, red = rumour
CONF = {"FFC6EFCE": "c", "FFFFEB9C": "t", "FFFFC7CE": "r"}
TIERS = {0: "main", 1: "epd", 2: "sub", 3: "aff", 4: "col"}


def fill(cell):
    f = cell.fill
    return f.fgColor.rgb if f is not None and f.fill_type else None


def num(v):
    """Return (value, approx) for a headcount cell; None when blank or '-'."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return [round(v), abs(v - round(v)) > 1e-9]
    s = str(v).strip()
    m = re.fullmatch(r"([<>~]?)\s*(\d+)", s)
    if m:
        return [int(m.group(2)), bool(m.group(1))]
    return None


def split_game(s):
    if not s or str(s).strip() in ("-", ""):
        return None
    s = str(s).replace("​", "").replace("´", "'").strip()
    m = re.match(r"^(.*?)\s*\((\d{4})\)\s*$", s)
    return [m.group(1).strip(), int(m.group(2))] if m else [s, None]


def simple_formula(ws, f):
    """Value of a formula like =J13*0.85 when the file has no cached result (e.g. saved by openpyxl)."""
    m = re.fullmatch(r"=\s*([A-Z]+)(\d+)\s*([*/+-])\s*([\d.]+)\s*", f)
    if not m:
        return None
    base = ws[m.group(1) + m.group(2)].value
    if not isinstance(base, (int, float)):
        return None
    k = float(m.group(4))
    return {"*": base * k, "/": base / k, "+": base + k, "-": base - k}[m.group(3)]


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def main():
    date = datetime.date.today().isoformat()
    if "--date" in sys.argv:
        date = sys.argv[sys.argv.index("--date") + 1]
    extra = json.loads(EXTRA.read_text(encoding="utf-8"))
    wb = openpyxl.load_workbook(XLSX)
    ws = wb["Nintendo"]  # formatting (confidence colours, links)
    wv = openpyxl.load_workbook(XLSX, data_only=True)["Nintendo"]  # cached values of formula cells

    # Headcount columns run from E up to the "Source" column; the side tables sit 7 columns after it
    srccol = next(c for c in range(5, 40) if str(ws.cell(2, c).value or "").strip() == "Source")
    ecols = range(5, srccol)
    SN, SD, SG = srccol + 7, srccol + 8, srccol + 9
    head = [ws.cell(2, c).value for c in ecols]
    years = [re.sub(r"^Employees \((.*)\)$", r"\1", str(h)) for h in head]

    # Side table, collaborators: name (S) -> last game (T)
    collab = {}
    for r in range(21, 60):
        n, g = ws.cell(r, SN).value, ws.cell(r, SD).value
        if n and n not in ("Collaborator",):
            collab[norm(n)] = split_game(g)

    studios, tier = [], None
    for r in range(3, 60):
        name = ws.cell(r, 3).value
        if not name:
            break
        t = ws.cell(r, 2).value
        if isinstance(t, (int, float)):
            tier = int(t)
        flag = (ws.cell(r, 1).value or "").strip()
        emp = []
        for i, c in enumerate(ecols):
            raw = wv.cell(r, c).value
            if raw is None and str(ws.cell(r, c).value or "").startswith("="):
                raw = simple_formula(ws, str(ws.cell(r, c).value))
            v = num(raw)
            formula = str(ws.cell(r, c).value or "").startswith("=")  # estimate computed in the sheet
            emp.append([v[0], CONF.get(fill(ws.cell(r, c)), "t"), v[1] or formula] if v else None)
        src = ws.cell(r, srccol).hyperlink.target if ws.cell(r, srccol).hyperlink else ws.cell(r, srccol).value
        if src and not str(src).startswith("http"):
            src = None
        x = extra["studios"].get(name, {})
        last = collab.get(norm(name))
        d = {
            "k": name, "n": x.get("disp", name), "t": TIERS[tier], "f": "new" if "New" in flag else ("fade" if "Dec" in flag else ""),
            "asof": (ws.cell(r, 4).value or "").replace("-", "").strip(), "e": emp, "src": src,
        }
        for key in ("r", "loc", "y", "own", "ot", "who", "known", "web", "note", "upcoming", "hide"):
            if x.get(key) is not None:
                d[key] = x[key]
        if last:
            d["g"] = last + [x.get("ms")]
        elif x.get("recent"):
            d["g"] = x["recent"]
            d["gx"] = 1  # not from the sheet
        if x.get("src"):
            d["s2"] = x["src"]
        studios.append(d)

    # EPD production groups (S/T/U rows 4-15)
    groups = []
    for r in range(4, 16):
        n = ws.cell(r, SN).value
        if not n:
            continue
        gx = extra["groups"].get(n, {})
        g = {"n": n, "d": (ws.cell(r, SD).value or "").strip(), "g": split_game(ws.cell(r, SG).value)}
        if gx.get("former"):
            g["former"] = 1
        if g["g"]:
            g["g"].append(gx.get("ms"))
        if gx.get("src"):
            g["src"] = gx["src"]
        if gx.get("note"):
            g["note"] = gx["note"]
        groups.append(g)

    # One card per EPD production group (PG#8 already has its own row: "∟PG #8 (Tokyo)")
    tokyo = next((x for x in studios if "PG #8" in x["k"]), None)
    for g in groups:
        if g.get("former"):
            continue
        gx = extra["groups"].get(g["n"], {})
        if g["n"] == "PG#8" and tokyo:
            if g["g"] and "g" not in tokyo:
                tokyo["g"] = g["g"]
            if gx.get("src"):
                tokyo.setdefault("s2", []).append(gx["src"])
            continue
        card = {"k": g["n"], "n": gx.get("disp", "Nintendo EPD " + g["n"]), "t": "epd", "f": "", "asof": "",
                "e": [None] * len(years), "src": None, "r": "jp", "loc": "Kyoto", "pg": 1,
                "ot": "Internal", "own": "Internal division of Nintendo Co., Ltd.",
                "note": g["d"] + "." + (" " + g["note"] if g.get("note") else "") +
                        " EPD doesn't publish staff per group; the whole Kyoto division is estimated at about 700 people."}
        for key in ("known", "who"):
            if gx.get(key):
                card[key] = gx[key]
        if g["g"]:
            card["g"] = g["g"]
        if gx.get("src"):
            card["s2"] = [gx["src"]]
        studios.append(card)

    # Development thread (data/thread.json, made by scripts/parse_thread.py): people, work and game credits
    GA = []
    tj = ROOT / "data" / "thread.json"
    if tj.exists():
        th = json.loads(tj.read_text(encoding="utf-8"))
        bykey = {x["k"]: x for x in studios}
        alias = {"PG#8": tokyo["k"] if tokyo else "PG#8", "EPD #1": "CoPG (former EPD#1+2)", "EPD #2": "CoPG (former EPD#1+2)"}
        tr = lambda k: alias.get(k, k)
        pf = extra.get("people_fixes", {})
        def fixp(txt):
            for k, v in pf.items():
                txt = re.sub(re.escape(k) + r"(?! \()", v, txt)
            return txt
        def put(card, people=None, rel=None, cur=None, about=None, wn=None):
            if people: card["ppl"] = [fixp(x) for x in people]
            if rel: card["rel"] = list(dict.fromkeys(rel))
            if cur: card["cur"] = list(dict.fromkeys(cur))
            if about: card["about"] = about
            if wn: card["wn"] = wn
        for code, g in th["groups"].items():
            card = bykey.get(tr(code))
            if not card or code == "EPD #1":
                continue
            ppl = []
            if g.get("producers"): ppl.append("Producers: " + ", ".join(g["producers"]))
            if g.get("leads"): ppl.append("Project leads: " + ", ".join(g["leads"]))
            series = ", ".join(g.get("series") or [])
            if series and re.match(r"PG#(?!8)\d+$", code):
                card["known"] = series  # the thread's series list is the best "known for" for a production group
                series = None
            put(card, ppl, g.get("recent"), g.get("current"), wn=series or None)
        kyoto = next((x for x in studios if x["k"].startswith("Nintendo PG #3")), None)
        if kyoto and th.get("management"):
            put(kyoto, [f"{k}: {', '.join(v)}" for k, v in th["management"].items()], about=th.get("epd_about"))
        for name, b in th["studios"].items():
            card = bykey.get(b.get("key") or "")
            if card:
                put(card, b.get("people"), b.get("releases"), b.get("current"), b.get("about"))
        for key, works in th["collab"].items():
            if key in bykey:
                put(bykey[key], wn=", ".join(works))
        for f in th["former"]:
            if f.get("key") in bykey:
                c = bykey[f["key"]]
                c["f"] = "fade"
                c["former"] = 1
        MC = extra.get("mc", {})
        for g in th["games"]:
            m = MC.get(g["g"] + "|" + g["p"]) or MC.get(g["g"]) or [None, None]
            GA.append([g["g"], g["p"], g["y"] or (m[2] if len(m) > 2 else None),
                       list(dict.fromkeys(tr(k) for k in g["k"])), g["o"],
                       extra.get("dates", {}).get(g["g"]), m[0], m[1]])

    # Credit corrections checked against end credits (data/extra.json -> credit_fixes)
    GROUPS = {"PG#3", "PG#4", "PG#5", "PG#6", "PG#7", "PG#9", "PG#10", "SDPG", "CoPG (former EPD#1+2)"}
    if tokyo:
        GROUPS.add(tokyo["k"])
    for g in GA:
        fx = extra.get("credit_fixes", {}).get(g[0])
        if not fx:
            continue
        keys = g[3]
        if "set_groups" in fx:
            keys = fx["set_groups"] + [k for k in keys if k not in GROUPS]
        g[3] = list(dict.fromkeys(keys + fx.get("add", [])))
        g[4] = list(dict.fromkeys(g[4] + fx.get("other", [])))
        if fx.get("src") and len(g) == 6:
            g.append(fx["src"])

    # Type legend ("Type / Description" block under the main table)
    legend = {}
    for r in range(3, 120):
        if str(ws.cell(r, 2).value).strip() == "Type" and str(ws.cell(r, 3).value).strip() == "Description":
            for rr in range(r + 1, r + 12):
                t, desc = ws.cell(rr, 2).value, ws.cell(rr, 3).value
                if not isinstance(t, (int, float)) or not desc:
                    break
                legend[TIERS[int(t)]] = str(desc).replace("\u200b", "").strip()
            break

    # Game credits (sheet "Hoja 3")
    credits = []
    if "Hoja 3" in wb.sheetnames:
        h = wb["Hoja 3"]
        for row in h.iter_rows(min_row=3, values_only=True):
            if row[1]:
                credits.append({"g": row[1], "p": row[2], "y": int(row[3]) if row[3] else None,
                                "main": row[4], "help": [v for v in row[5:9] if v]})

    data = {"date": date, "years": years, "S": studios, "G": groups, "C": credits, "TL": legend, "GA": GA, "SC": {k: v for k, v in extra.get("scores", {}).items() if not k.startswith("_")}}
    js = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = TPL.read_text(encoding="utf-8")
    d = datetime.date.fromisoformat(date)
    html = (html.replace("/*DATA*/null", js).replace("{{DATE_ISO}}", date)
                .replace("{{DATE}}", f"{d.day} {d.strftime('%b %Y')}").replace("{{SSR}}", ssr(studios)))
    out = OUT
    if "--preview" in sys.argv:
        # Standalone preview page: fonts from Google Fonts, no <html>/<head> wrapper
        out = pathlib.Path(sys.argv[sys.argv.index("--preview") + 1])
        html = html.replace("{{HEAD}}", '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
                            '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Fredoka:wght@700&family=IBM+Plex+Sans:wght@400;500;600&display=swap">')
    else:
        fonts = (ROOT / "src" / "fonts.css").read_text(encoding="utf-8")
        title = "Nintendo Party Superstars · Who makes Nintendo's games"
        desc = "Nintendo's in-house EPD teams, subsidiaries, affiliates and frequent partner studios: headcount since 2022, owners, people and the Metascore of each one's latest game."
        ld = {"@context": "https://schema.org", "@graph": [
            {"@type": "WebSite", "@id": SITE + "#site", "url": SITE, "name": "Nintendo Party Superstars", "inLanguage": "en"},
            {"@type": "Dataset", "name": "Nintendo's development studios", "description": desc, "url": SITE,
             "dateModified": date, "license": "https://creativecommons.org/licenses/by/4.0/", "isAccessibleForFree": True,
             "creator": {"@type": "Person", "name": "Nakuto", "url": BSKY, "sameAs": [BSKY]},
             "keywords": ["Nintendo", "Nintendo EPD", "game developers", "second-party studios", "headcount", "Metascore"],
             "about": [{"@type": "Organization", "name": s["n"]} for s in studios if not s.get("pg")]}]}
        e = html_escape
        head = (f'<link rel="canonical" href="{SITE}"><meta name="theme-color" content="#f5f0e3">'
                '<meta name="robots" content="index,follow,max-image-preview:large">'
                '<link rel="icon" href="favicon.svg" type="image/svg+xml"><link rel="apple-touch-icon" href="apple-touch-icon.png">'
                f'<meta property="og:type" content="website"><meta property="og:site_name" content="Nintendo Party Superstars"><meta property="og:url" content="{SITE}">'
                f'<meta property="og:title" content="{e(title)}"><meta property="og:description" content="{e(desc)}"><meta property="og:locale" content="en_GB">'
                f'<meta property="og:image" content="{SITE}og.png"><meta property="og:image:width" content="1200"><meta property="og:image:height" content="630">'
                '<meta property="og:image:alt" content="Nintendo Party Superstars: who makes Nintendo\'s games">'
                f'<meta name="twitter:card" content="summary_large_image"><meta name="author" content="Nakuto">'
                '<link rel="preload" href="f/fredoka-700-latin.woff2" as="font" type="font/woff2" crossorigin>'
                '<link rel="preload" href="f/plex-sans-400-latin.woff2" as="font" type="font/woff2" crossorigin>'
                f'<style>{fonts}</style>'
                '<script type="application/ld+json">' + json.dumps(ld, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/") + "</script>")
        html = html.replace("{{HEAD}}", head)
        i = html.index('<div class="wrap">')
        html = ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
                '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
                + html[:i] + "</head><body>" + html[i:] + "</body></html>")
        (ROOT / "public" / "sitemap.xml").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'
            f'<url><loc>{SITE}</loc><lastmod>{date}</lastmod><changefreq>weekly</changefreq></url></urlset>\n', encoding="utf-8")
        (ROOT / "public" / "robots.txt").write_text(f"User-agent: *\nAllow: /\n\nSitemap: {SITE}sitemap.xml\n", encoding="utf-8")
    raw = len(html.encode())
    if "--no-min" not in sys.argv:
        html = minify(html)
    out.parent.mkdir(exist_ok=True)
    out.write_text(html, encoding="utf-8")
    size = len(html.encode())
    print(f"built {out}: {len(studios)} studios, {len(groups)} EPD groups, {len(credits)} credits, "
          f"{raw:,} -> {size:,} bytes ({len(gzip.compress(html.encode(), 9)):,} gzipped)")


def ssr(studios):
    """Plain studio list baked into the page so search engines and no-JS readers see the content; the script replaces it."""
    e = html_escape
    out = []
    for s in studios:
        if s["t"] == "main" or s.get("hide") or s["k"].startswith("Nintendo PG #3-7"):  # same rule as SHOWN() in the template
            continue
        meta = " · ".join(str(x) for x in (s.get("loc"), s.get("y")) if x)
        known = f'<p class="known"><b>Known for</b>{e(s["known"])}</p>' if s.get("known") else ""
        out.append(f'<article class="card"><div class="band"><h3>{e(s["n"])}</h3><div class="meta"><span>{e(meta)}</span></div></div>{known}</article>')
    return "".join(out)


def mincss(c):
    c = re.sub(r"/\*.*?\*/", "", c, flags=re.S)
    c = re.sub(r"\s+", " ", c)
    c = re.sub(r"\s*([{};,>])\s*", r"\1", c)
    c = re.sub(r":\s+", ":", c)
    return c.replace(";}", "}").strip()


def minjs(js):
    """Minify with terser when it is installed (npm install), otherwise leave the script as it is."""
    bin_ = ROOT / "node_modules" / ".bin" / "terser"
    if not bin_.exists():
        print("  (terser not installed: script left unminified; run npm install)")
        return js
    r = subprocess.run([str(bin_), "--ecma", "2020", "-c", "passes=2", "-m", "--toplevel"],
                       input=js, capture_output=True, text=True, encoding="utf-8")
    if r.returncode:
        sys.exit("terser failed:\n" + r.stderr)
    return r.stdout.strip().rstrip(";")


def minify(html):
    parts = re.split(r"(<style>.*?</style>|<script>.*?</script>|<script type=\"application/ld\+json\">.*?</script>)", html, flags=re.S)
    res = []
    for p in parts:
        if p.startswith("<style>"):
            res.append("<style>" + mincss(p[7:-8]) + "</style>")
        elif p.startswith("<script>"):
            res.append("<script>" + minjs(p[8:-9]) + "</script>")
        elif p.startswith("<script type="):
            res.append(p)
        else:
            p = re.sub(r"<!--.*?-->", "", p, flags=re.S)
            p = re.sub(r">\s*\n\s*<", "><", p)
            res.append(re.sub(r"\n\s*", " ", p).strip())
    return "".join(res)

if __name__ == "__main__":
    main()
