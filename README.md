# Nintendo Party Superstars

Who makes Nintendo's games: EPD production groups, subsidiaries, affiliates and frequent partner studios, with headcount since 2022, owners and the Metascore of each one's latest game.

Sister project of [Soft Reset](https://softreset.dev). Data by Nakuto, CC BY 4.0.

## How it's built

| File | What it is |
|---|---|
| `data/Nintendo.xlsx` | The tracking sheet. **Edit this** to update headcounts, add studios or change last games. Cell colours matter: green = company figure, yellow = third party, red = rumour. Column A: `(New)` or `(Dec?)`. |
| `data/extra.json` | Everything that isn't in the sheet: founding year, city, owner, people, known for, website, Metascore, notes, extra sources. Keys = the sheet's `Name` column, exactly. |
| `data/thread.txt` | The first-party development thread pasted as plain text. Paste a newer copy over it and run `python3 scripts/parse_thread.py` to refresh people, recent/current work and game credits (`data/thread.json`). |
| `src/template.html` | Page design and code. |
| `scripts/build.py` | Reads the data files and writes `public/index.html`, `sitemap.xml` and `robots.txt`. |

```bash
pip install openpyxl
npm install                          # terser, to minify the page script
python3 scripts/build.py             # -> public/index.html (minified), sitemap.xml, robots.txt
python3 scripts/build.py --no-min    # readable output, for debugging
```

### SEO and weight
- The build writes the title, description, canonical URL, Open Graph/Twitter tags, JSON-LD (`WebSite` + `Dataset`, CC BY 4.0, by Nakuto), `sitemap.xml` and `robots.txt`. **Change `SITE` in `scripts/build.py` once the domain is decided.**
- The studio list is baked into the HTML (readable without JavaScript and by search engines); the script replaces it with the full cards.
- `public/og.png` is the share image (1200×630), made from `scripts/og.html` with a headless-browser screenshot.
- CSS and HTML are minified by the build, the script by terser. The "Every game" table is only drawn when it is opened.
- Fonts are self-hosted, without hinting, Bricolage pinned to one optical size; only Fredoka 700 and Plex Sans 400 are preloaded.

### Nintendo's own headcount updates itself
Nintendo publishes its headcount every quarter with its financial results, on its [company profile](https://www.nintendo.co.jp/corporate/en/outline/index.html). Once a day the GitHub Action runs `scripts/update_nintendo.py`, which reads that page and, when it shows a new quarter, adds an `Employees (MM/YYYY)` column before `Source` with the group and Nintendo Co., Ltd. figures, then rebuilds and redeploys. If Nintendo changes the page layout the job fails with a clear message instead of writing wrong numbers.

## Publishing (GitHub Pages)

The site is published from `public/` by the GitHub Action in `.github/workflows/build.yml`. Pushing changes in `data/`, `src/` or `scripts/` to `main` rebuilds and republishes it; you can also run it by hand from **Actions → Build and publish → Run workflow**.

One-time setup: **Settings → Pages → Build and deployment → Source: GitHub Actions**.

All asset paths are relative, so the site works both at `https://<user>.github.io/<repo>/` and on a custom domain. The canonical/OG/sitemap address comes from the Pages URL automatically (`SITE_URL`). To use a custom domain, add it under **Settings → Pages → Custom domain** and run the workflow again.

## Adding a studio
1. Add a row to the sheet in the right block (type 2 subsidiary, 3 affiliate, 4 collaborator), with its colours.
2. For collaborators, add the name + last game in the side table (columns S/T).
3. Add an entry with the same name to `data/extra.json`.
