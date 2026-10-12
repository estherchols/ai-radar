# AI Radar

A weekly AI industry dashboard built around one idea. Value accrues to whoever controls the scarcest input, and right now there are three. Compute, capital and energy. Rules that steer money are tracked under capital. Every item on the page is tagged by the constraint it moves and the direction it moves it.

## Files

| File | What it is |
| --- | --- |
| `index.html` | The page. Drawing code, styles and a built-in copy of the data. Open it directly or host it anywhere static. |
| `data.json` | Every number and every sentence on the page. This is the file you edit. |
| `update.py` | Refreshes GPU and storage prices in `data.json` through the Firecrawl API. |
| `fixtures/` | Saved Firecrawl-shaped responses so you can test the updater without a key or a network call. |
| `data.json.bak` | Written by `update.py` before each save. One level of undo. |

## How the page loads its data

The page asks for `data.json` first. If that request fails, or if you open the file straight from disk, it falls back to the copy built into `index.html`. Both paths render the same page. The masthead under the header shows the issue number, the issue date and the date prices were last checked.

`update.py` keeps the built-in copy in step with `data.json` by itself. If you edit `data.json` by hand, run this once afterwards.

```
python3 airadar/update.py --sync-only
```

## Running the price updater

You need Python 3 and the `requests` package, plus a Firecrawl API key.

```
pip install requests
export FIRECRAWL_API_KEY=fc-your-key
python3 airadar/update.py --dry-run     # prints what would change, writes nothing
python3 airadar/update.py               # writes data.json and data.json.bak
```

Without a key the script stops at once with exit code 2 and tells you why.

What it does for each entry in `price_sources`.

1. Sends one request to `POST https://api.firecrawl.dev/v2/scrape` with a JSON extraction schema and a plain-language prompt. Cached copies are skipped by setting `maxAge` to 0.
2. Converts instance prices to a per GPU-hour rate by dividing by the GPU count.
3. Matches each extracted row to a row in `data.json` by provider, GPU model and the SKU tokens in `match_sku`.
4. Ignores moves of 0.5 percent or less. For anything larger it updates the price, adds a "Previously $X" note to the row, and adds a dated point to the price history chart with the source link.
5. Holds back any single move above 50 percent for a person to look at, since that usually means the page layout changed. Change the limit with `--max-change`.
6. Sets the "prices checked" date, adds an entry to the changelog and refreshes the built-in copy in `index.html`.

Safety rules built into the script.

- It never deletes a row.
- It never touches signals (including their per-audience lines), the weekly headline, the audience implications, the market summaries, deals, policies, takes, the constraint board, layers, models, the about block or the methodology. If any of those would change, it refuses to save.
- A failed source is reported and skipped. Everything else still updates, and the exit code is 1.
- Entries it cannot match are listed and left alone. Add a row and a `match_sku` token to `data.json` if you want them tracked.

### Testing without Firecrawl

```
python3 airadar/update.py --fixture fixtures --dry-run --today 2026-10-18
python3 airadar/update.py --fixture fixtures --today 2026-10-18
```

The fixtures hold hand-made responses in the shape Firecrawl returns. They include three price moves, one entry with no matching row and one source with no file. Restore the original with `cp data.json.bak data.json` afterwards.

### Options

| Flag | Effect |
| --- | --- |
| `--dry-run` | Print the diff and write nothing |
| `--fixture PATH` | Read responses from a folder (one `<source id>.json` each) and skip the network |
| `--only ID ...` | Check only the listed source ids |
| `--today YYYY-MM-DD` | Override the date, useful with fixtures |
| `--max-change 0.5` | Hold back any single move larger than this fraction |
| `--no-inline` | Leave the built-in copy in `index.html` alone |
| `--sync-only` | Only refresh the built-in copy, no key needed |

## What is automated and what is editorial

Automated by `update.py`.

- GPU rental prices and storage prices in the compute tables
- The price history chart points for H100, H200 and B200
- The "prices checked" date, the compute "as of" date and the changelog entry for price moves

Editorial, written by hand in `data.json`.

- The weekly `headline`, the three `implications` per audience (investors, bigtech, builders) and the `markets_sowhat` summaries
- Signals, including the constraint tags, the per-audience `audience` lines, the Moves line and the sources
- The thesis, the constraint board and the three takes
- Capital deals, model entries and policy items
- The ecosystem layers, about, methodology and notes
- The issue number and update date in `meta`

The Silicon Data neocloud index line on the chart and the chart annotations are also hand maintained.

## Weekly routine

1. Run `python3 airadar/update.py --dry-run` and read the diff.
2. Run `python3 airadar/update.py` to save it. Open any source link behind a large move and confirm the number on the page.
3. Add new signals, deals, policy items and model entries by hand. Every figure needs a date and a link. Leave out anything you cannot source.
4. Revisit the board. Say which constraint moved and why, and update the "what would flip this" line.
   Rewrite the headline, the nine implications and the audience line on each new signal. Every implication lists `evidence` ids, which must match an `id` on a signal, deal, model or policy, or `price-<row id>` for a price row. `{lo_row}` and `{hi_row}` point at the cheapest and dearest H100 rows.
5. Raise `meta.issue`, set `meta.updated` and `meta.next_issue`, and add a changelog entry for the editorial changes.
6. Run `python3 airadar/update.py --sync-only` so the built-in copy matches.
7. Open `index.html` once from disk and once over http to confirm both load paths look the same, then publish.

## Notes on the Firecrawl request

Reference pages are https://docs.firecrawl.dev/api-reference/endpoint/scrape and https://docs.firecrawl.dev/features/llm-extract.

The request sends `url`, `formats` containing one object with `type` set to `json` plus `schema` and `prompt`, `maxAge` set to 0 and `timeout` in milliseconds. The key goes in an `Authorization: Bearer` header. The reply is read from `data.json` inside the response body, and the script checks `success` first. Retries with a short wait are used for rate limit and server errors.

Secondary sources, meaning the Holori, Spheron and Cloudprice mirror pages used for AWS and Azure, are tagged as secondary on every history point they create. Replace them with official price list pages when you find a reliable one.
