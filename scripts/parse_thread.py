#!/usr/bin/env python3
"""Turn data/thread.txt (the first-party development thread, pasted as plain text)
into data/thread.json: people, series, recent/current work per EPD group and studio,
plus every game with its credits.

Usage:  python3 scripts/parse_thread.py
Paste a newer copy of the thread over data/thread.txt and run it again.
Unknown studio names are kept as plain text ("other" credits) and listed at the end.
"""
import json, pathlib, re

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "thread.txt"
OUT = ROOT / "data" / "thread.json"
ZW = "​"

# Thread name -> key used in Nintendo.xlsx (column C) or EPD group code (side table)
STUDIO_BLOCKS = {
    "Monolith Soft": "MONOLITH SOFTWARE", "Retro Studios": "Retro Studios", "Nintendo Cube": "Nintendo Cube",
    "1-UP Studio": "1-Up Studio", "Next Level Games": "Next Level Games",
    "Nintendo Software Technology": "Nintendo Software Technology (NST)",
    "Nintendo European Research & Development": "Nintendo European Research & Development (N.E.R.D.)",
    "iQue": "iQue", "SRD": "SRD", "Nintendo Pictures": "Nintendo Pictures",
    "Shiver Entertainment": "Shiver Entertiment", "Nintendo Studios Singapore": "Nintendo Studios Singapore",
    "Game Freak": "Game Freak", "HAL Laboratory": "HAL Laboratory", "Intelligent Systems": "Intelligent Systems",
    "Grezzo": "Grezzo", "Genius Sonority": "Genious Sonority", "The Pokémon Works": None,
}
COLLAB = {
    "acquire": "ACQUIRE Corp.", "arika": "Arika", "artepiazza": "ArtePiazza", "camelot": "Camelot Software Planning",
    "creatures": "Creatures Inc.", "eighting": "Eighting", "forever entertainment": "Forever Entertainment",
    "good feel": "Good-Feel", "good-feel": "Good-Feel", "indieszero": "indieszero", "mages": "Mages Inc.",
    "mercurysteam": "MercurySteam", "platinumgames": "PlatinumGames", "platinum games": "PlatinumGames",
    "spike chunsoft": "Spike Chunsoft", "tantalus": "Tantalus Media", "arzest": "Arzest", "syn sophia": "Syn Sophia",
}
# Aliases used inside game credits (lowercase, matched as whole words)
CREDIT_ALIASES = dict(COLLAB, **{
    "monolithsoft": "MONOLITH SOFTWARE", "monolift soft": "MONOLITH SOFTWARE", "monolith soft": "MONOLITH SOFTWARE",
    "nintendo software technology": "Nintendo Software Technology (NST)", "nst": "Nintendo Software Technology (NST)",
    "nd cube": "Nintendo Cube", "ndcube": "Nintendo Cube", "nintendo cube": "Nintendo Cube",
    "1-up studio": "1-Up Studio", "ique": "iQue", "nerd": "Nintendo European Research & Development (N.E.R.D.)",
    "retro studios": "Retro Studios", "next level games": "Next Level Games", "hal laboratory": "HAL Laboratory",
    "intelligent systems": "Intelligent Systems", "game freak": "Game Freak", "gamefreak": "Game Freak",
    "grezzo": "Grezzo", "ilca": "ILCA", "genius sonority": "Genious Sonority", "shiver": "Shiver Entertiment",
    "mages inc.": "Mages Inc.", "tantalus media": "Tantalus Media", "camelot software": "Camelot Software Planning",
    # every Bandai Namco spelling in the thread goes to its development arm; the Singapore studio is Nintendo's now
    "bandai namco": "Bandai Namco Studios", "bandai namco games": "Bandai Namco Studios", "namco-bandai games": "Bandai Namco Studios",
    "bandai namco studios": "Bandai Namco Studios", "bandai namco studio s": "Bandai Namco Studios", "bandai namco studios malaysia": "Bandai Namco Studios",
    "bandai namco studios singapore": "Nintendo Studios Singapore",
})
GROUP_HEAD = [
    (r"Production Group No\. 1\b", "EPD #1"), (r"Co-Production Group", "CoPG (former EPD#1+2)"),
    (r"Smart Device Production Group", "SDPG"), (r"Production Group No\. (\d+)", None),
]
# Known typos / outdated titles in the thread -> corrected title
TITLE_FIXES = {
    "Luigi Mansion 2 HD": "Luigi's Mansion 2 HD",
    "Super Mario Bros. Wonder: Nintendo Switch 2 Edition + Meetup in Bellalel Park": "Super Mario Bros. Wonder: Nintendo Switch 2 Edition + Meetup in Bellabel Park",
    "Pikmin 4: Nintendo Switch Edition + Dandori Academy": "Pikmin 4: Nintendo Switch 2 Edition + Dandori Academy",
    "Xenoblade Chronicles: Definitive Edition -Nintendo Switch 2 Edition": "Xenoblade Chronicles: Definitive Edition - Nintendo Switch 2 Edition",
    "Hyrule Warriors: Age of Calamity": "Hyrule Warriors: Age of Calamity – Definitive Edition",  # only fixed in the 2027 block, see below
    "Pokémon Brilliant Diamond/Shining Pearl": "Pokémon Brilliant Diamond / Shining Pearl",
    "Baiten Kaitos I + II HD Remaster": "Baten Kaitos I & II HD Remaster",
    "Super Mario Party Jamboree – Nintendo Switch 2 Edition ＋ Jamboree TV": "Super Mario Party Jamboree – Nintendo Switch 2 Edition + Jamboree TV",
}
KV = re.compile(r"^(Manager / Producer|Project Leads|Series|Recent Titles|Recent|Current|President|Managing Director|"
                r"General Producer|Releases|Projects)\s*[-:]\s*(.+)$")


def items(v):
    v = v.replace(ZW, "").strip().rstrip(",")
    parts, depth, cur = [], 0, ""
    for ch in v:  # split on commas outside parentheses
        depth += ch == "("; depth -= ch == ")"
        if ch == "," and depth == 0:
            parts.append(cur.strip()); cur = ""
        else:
            cur += ch
    parts.append(cur.strip())
    return [p for p in parts if p and p.upper() != "TBA"]


def group_code(t):
    t = t.lower().replace(" ", "")
    if "co-production" in t or "coproduction" in t:
        return "CoPG (former EPD#1+2)"
    if re.search(r"epd(pgsd|sg)$", t):
        return "SDPG"
    m = re.search(r"epd(?:pgno\.?)?(\d+)", t)
    if m:
        n = int(m.group(1))
        return {1: "EPD #1", 2: "EPD #2"}.get(n, f"PG#{n}")
    return None


def credit_keys(credits):
    keys, other, last_group = [], [], False
    for raw in credits:
        tok = raw.strip().rstrip("?").strip()
        if not tok or tok.lower() == "nintendo":
            continue
        if re.fullmatch(r"\d+", tok) and last_group:  # "EPD 4/7"
            keys.append(f"PG#{tok}"); continue
        for t in re.split(r"\s+or\s+", tok):
            if re.fullmatch(r"\d+", t.strip()) and last_group:
                keys.append(f"PG#{t.strip()}"); continue
            g = group_code(t)
            if g:
                keys.append(g); last_group = True; continue
            last_group = False
            low = t.lower().strip(" ?")
            hit = next((v for a, v in sorted(CREDIT_ALIASES.items(), key=lambda x: -len(x[0])) if re.search(r"(^|\W)" + re.escape(a) + r"($|\W)", low)), None)
            if hit:
                keys.append(hit)
            elif "epd" not in low:
                other.append(t.strip())
    return list(dict.fromkeys(keys)), other


def split_title(line):
    if not line.endswith(")"):
        return None
    depth = 0
    for i in range(len(line) - 1, -1, -1):
        depth += line[i] == ")"; depth -= line[i] == "("
        if depth == 0:
            return line[:i].strip(), line[i + 1:-1]
    return None


def main():
    lines = [l.replace(ZW, "").rstrip() for l in SRC.read_text(encoding="utf-8").splitlines()]
    out = {"groups": {}, "studios": {}, "collab": {}, "former": [], "defunct": [], "games": [], "unmapped": [],
           "epd_about": "", "management": {}}
    sec, cur = None, None
    plat, year = None, None
    for i, l in enumerate(lines):
        s = l.strip()
        if not s:
            continue
        # section switches
        if s == "Nintendo Entertainment Planning & Development": sec = "epd"; continue
        if s == "Nintendo EPD Management": sec = "mgmt"; continue
        if s in ("Nintendo Subsidiaries", "Nintendo Affiliates"): sec = "studios"; cur = None; continue
        if s == "Frequent Collaboration Studios": sec = "collab"; continue
        if s.startswith("Major Publisher Collaboration") or s.startswith("First Party-IP"): sec = "skip"; continue
        if s == "Former Collaborators": sec = "former"; continue
        if s == "Defunct Collaborators": sec = "defunct"; continue
        if s == "Games": sec = "games"; continue
        if s.startswith("Other threads"): break

        if sec == "epd":
            out["epd_about"] = s; continue
        if sec in ("mgmt", "groups"):
            m = re.match(r"^(General Manager|Deputy General Managers|Executive Officer|Senior Officer)\s*-\s*(.+)$", s)
            if m:
                out["management"][m.group(1)] = items(m.group(2)); continue
            if s.startswith("Nintendo EPD"):
                sec = "groups"; cur = None
                for pat, code in GROUP_HEAD:
                    m = re.search(pat, s)
                    if m:
                        cur = code or f"PG#{m.group(1)}"; break
                if cur:
                    out["groups"].setdefault(cur, {"head": s})
                continue
            m = KV.match(s)
            if m and cur:
                k = {"Recent Titles": "recent", "Recent": "recent", "Current": "current", "Series": "series",
                     "Manager / Producer": "producers", "Project Leads": "leads"}.get(m.group(1), m.group(1))
                out["groups"][cur][k] = items(m.group(2))
            continue
        if sec == "studios":
            if s in STUDIO_BLOCKS:
                cur = s
                out["studios"][cur] = {"key": STUDIO_BLOCKS[s]}
                continue
            if cur is None:
                continue
            st = out["studios"][cur]
            m = KV.match(s)
            if m:
                k = m.group(1)
                if k in ("Releases", "Projects"):
                    st.setdefault("releases", []).extend(items(re.sub(r"^Production Group 1 \([^)]*\)\s*-\s*", "", m.group(2))))
                elif k == "Current":
                    st.setdefault("current", []).extend(items(m.group(2)))
                else:
                    st.setdefault("people", []).append(f"{m.group(2).strip()} ({k.lower()})" if "(" not in m.group(2) else m.group(2).strip())
                continue
            m = re.match(r"^Production Group 2\s*-\s*(.+)$", s)
            if m:
                st.setdefault("releases", []).extend(items(m.group(1))); continue
            m = re.match(r"^Kyoto Group \(([^)]+)\)", s)
            if m:
                st.setdefault("people", []).append(f"{m.group(1)} (Kyoto group)"); continue
            if "Speculation Thread" in s:
                continue
            if "about" not in st:
                st["about"] = s
            continue
        if sec == "collab":
            m = re.match(r"^(.+?)\s+-\s+(.+)$", s)
            if m:
                name = m.group(1).strip(" ,").lower()
                key = next((v for a, v in COLLAB.items() if name.startswith(a)), None)
                if key:
                    out["collab"][key] = items(m.group(2))
                else:
                    out["unmapped"].append(m.group(1).strip(" ,"))
            continue
        if sec in ("former", "defunct"):
            name = re.split(r"\s+[-(]", s, maxsplit=1)[0].strip()
            key = next((v for a, v in COLLAB.items() if name.lower().startswith(a)), None)
            out[sec].append({"name": name, "key": key, "work": s[len(name):].strip(" -()")})
            continue
        if sec == "games":
            if re.match(r"^Nintendo Switch 2 \(Developed", s): plat, year = "Switch 2", None; continue
            if re.match(r"^Nintendo Switch \(Developed", s): plat, year = "Switch", None; continue
            if re.match(r"^Nintendo 3DS \(", s): plat, year = "3DS", None; continue
            if s == "Mobile": plat, year = "Mobile", None; continue
            m = re.match(r"^\[~+(\d{4})", s)
            if m:
                year = int(m.group(1)); continue
            st = split_title(s)
            if not st:
                continue
            title, cr = st
            title = title.strip()
            if title != "Hyrule Warriors: Age of Calamity" or (year or 0) >= 2027:
                title = TITLE_FIXES.get(title, title)
            keys, other = credit_keys(re.split(r"\s*/\s*", cr))
            if any(g["g"] == title and g["p"] == plat for g in out["games"]):
                continue  # listed twice in the thread
            out["games"].append({"g": title, "p": plat, "y": year, "k": keys, "o": other})
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    keys = {}
    for g in out["games"]:
        for k in g["k"]:
            keys[k] = keys.get(k, 0) + 1
    print(f"{len(out['groups'])} EPD groups, {len(out['studios'])} studio blocks, {len(out['collab'])} collaborators, {len(out['games'])} games")
    print("credits per key:", dict(sorted(keys.items(), key=lambda x: -x[1])))
    others = sorted({o for g in out["games"] for o in g["o"]})
    print("not in the sheet:", others)


if __name__ == "__main__":
    main()
