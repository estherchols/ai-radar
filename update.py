#!/usr/bin/env python3
"""Refresh the price data in data.json from the provider pricing pages.

Uses the Firecrawl v2 scrape endpoint with the JSON extraction format.
Only the compute tables, the price history and the changelog are touched.
Signals, deals, policies, theses, the constraint board, layers, models and
the about block are editorial and are checked for changes before any write.

Examples
  FIRECRAWL_API_KEY=fc-... python3 update.py --dry-run
  FIRECRAWL_API_KEY=fc-... python3 update.py
  python3 update.py --fixture fixtures --dry-run --today 2026-10-18
  python3 update.py --sync-only            (re-embed data.json into index.html)

Firecrawl references (verified 11 Oct 2026)
  https://docs.firecrawl.dev/api-reference/endpoint/scrape
  https://docs.firecrawl.dev/features/llm-extract
  https://docs.firecrawl.dev/features/fast-scraping
"""
import argparse
import copy
import datetime
import difflib
import json
import os
import re
import shutil
import sys
import time
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
API_URL = "https://api.firecrawl.dev/v2/scrape"
THRESHOLD = 0.005          # a matched rate must move by more than 0.5 percent to count as a change
HOURS_PER_MONTH = 730      # same conversion the page uses for hourly storage rates
PLAUSIBLE_PER_GPU_HOUR = (0.2, 60.0)
HISTORY_MODELS = ("H100", "H200", "B200")
PROTECTED = ("meta", "thesis", "how_to_read", "board", "signals", "takes", "layers", "models",
             "deals", "policies", "policy_timeline", "notes", "about", "methodology")
INLINE_RE = re.compile(r"(<!-- DATA-INLINE-START -->\s*)(.*?)(\s*<!-- DATA-INLINE-END -->)", re.S)


# ---------------------------------------------------------------- small helpers
def norm(s):
    """Lowercase and strip everything except letters and digits, so 'a3-highgpu-8g' matches 'A3 HighGPU 8G'."""
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def rel_diff(old, new):
    return abs(new - old) / old if old else float("inf")


def host(url):
    return urlparse(url).netloc.replace("www.", "")


def log(msg=""):
    print(msg, flush=True)


# ---------------------------------------------------------------- firecrawl
def build_schema(kinds):
    props = {}
    if "gpu" in kinds:
        props["gpu_prices"] = {
            "type": "array",
            "description": "One entry per on-demand GPU product or instance type listed on the page.",
            "items": {
                "type": "object",
                "properties": {
                    "gpu_model": {"type": "string", "description": "GPU model such as H100, H200 or B200."},
                    "sku_or_instance": {"type": "string", "description": "Exact product, SKU or instance type name as printed on the page."},
                    "price_usd": {"type": "number", "description": "Price in US dollars exactly as printed, without conversion."},
                    "billing_unit": {"type": "string", "description": "Either per_gpu_hour or per_instance_hour."},
                    "gpus_per_instance": {"type": "integer", "description": "Number of GPUs in the instance when the price is per instance."},
                    "conditions": {"type": "string", "description": "Region, commitment, tax or other conditions stated next to the price."},
                },
                "required": ["gpu_model", "sku_or_instance", "price_usd", "billing_unit"],
            },
        }
    if "storage" in kinds:
        props["storage_prices"] = {
            "type": "array",
            "description": "One entry per object storage product or class listed on the page.",
            "items": {
                "type": "object",
                "properties": {
                    "product": {"type": "string", "description": "Product, storage class and location as printed on the page."},
                    "price_per_gb_month": {"type": "number", "description": "Price in US dollars exactly as printed."},
                    "unit": {"type": "string", "description": "The unit printed on the page, for example per GB-month, per GiB-month or per GiB-hour."},
                },
                "required": ["product", "price_per_gb_month", "unit"],
            },
        }
    return {"type": "object", "properties": props}


def build_prompt(src, today):
    parts = [
        "Today is %s." % today,
        "Extract published list prices in US dollars from this page.",
        "Include only on-demand or pay-as-you-go rates. Ignore spot, preemptible, reserved, committed-use and cluster rates.",
        "Copy product and instance names exactly as printed. Do not estimate or invent any value.",
    ]
    if "gpu" in src["kinds"]:
        parts.append("For GPU compute return gpu_model, sku_or_instance, price_usd, billing_unit (per_gpu_hour or per_instance_hour) and gpus_per_instance when the price is per instance.")
    if "storage" in src["kinds"]:
        parts.append("For object storage return the standard storage class only, with product, price_per_gb_month and the printed unit.")
    if src.get("hint"):
        parts.append(src["hint"])
    return " ".join(parts)


def scrape(src, today, api_key, retries=3):
    """POST /v2/scrape with a json format. Returns the parsed response body."""
    import requests  # imported here so fixture runs need only the standard library
    body = {
        "url": src["url"],
        "formats": [{"type": "json", "schema": build_schema(src["kinds"]), "prompt": build_prompt(src, today)}],
        "maxAge": 0,            # force a fresh scrape, the default is a 2 day cache
        "timeout": 120000,      # milliseconds
    }
    headers = {"Authorization": "Bearer " + api_key, "Content-Type": "application/json"}
    last = None
    for attempt in range(retries + 1):
        try:
            r = requests.post(API_URL, headers=headers, json=body, timeout=(10, 180))
        except requests.RequestException as exc:
            last = "network error %s" % exc
        else:
            if r.status_code == 200:
                return r.json()
            last = "http %s %s" % (r.status_code, r.text[:200].replace("\n", " "))
            if r.status_code in (400, 401, 402, 403):   # not worth retrying
                break
        if attempt < retries:
            time.sleep(2 * (2 ** attempt))               # 2s, 4s, 8s for 429 and 5xx
    raise RuntimeError(last)


def extract(resp):
    """Pull data.json out of a Firecrawl response, raising on an unsuccessful one."""
    if not isinstance(resp, dict) or not resp.get("success", False):
        raise RuntimeError("response not successful")
    js = (resp.get("data") or {}).get("json")
    if not isinstance(js, dict):
        raise RuntimeError("response has no data.json object")
    return js


# ---------------------------------------------------------------- matching and updating
class Report:
    def __init__(self):
        self.changes = []      # human sentences for the changelog
        self.unchanged = 0
        self.unmatched = []
        self.warnings = []
        self.held = []
        self.errors = []


def strip_auto_note(cond):
    return re.sub(r"\s*Previously \$[0-9.,]+ before the \d{4}-\d{2}-\d{2} check\.", "", cond)


def history_series(data, row):
    ph = data["price_history"].setdefault(row["gpu"], {"series": [], "index": None, "annotations": []})
    for s in ph["series"]:
        if s.get("row_id") == row["id"]:
            return s
    s = {"provider": row["provider"], "row_id": row["id"], "points": []}
    if any(x["provider"] == row["provider"] for x in ph["series"]):
        s["name"] = row["label"]
    ph["series"].append(s)
    return s


def add_history_point(data, row, old_ph, new_ph, src, today, prev_checked):
    s = history_series(data, row)
    if not s["points"]:
        s["points"].append({
            "date": prev_checked or today, "price_per_gpu_hour_usd": old_ph, "type": row["source_type"],
            "note": "Baseline taken from the stored table before the first automated check.",
            "source_url": row["source_url"], "source_label": host(row["source_url"]) + ", stored value"})
    s["points"] = [p for p in s["points"] if p["date"] != today]       # one point per day
    s["points"].append({
        "date": today, "price_per_gpu_hour_usd": new_ph, "type": src["source_type"],
        "note": "Automated check of %s. Previously $%.2f per GPU-hour." % (host(src["url"]), old_ph),
        "source_url": src["url"], "source_label": "%s, checked %s" % (host(src["url"]), today)})
    s["points"].sort(key=lambda p: p["date"])
    slots = data["compute"].setdefault("provider_slots", {})
    if row["provider"] not in slots:
        free = [n for n in range(1, 7) if n not in slots.values()]
        if free:
            slots[row["provider"]] = free[0]


def apply_gpu(data, src, items, today, rep, max_change):
    rows = data["compute"]["gpu"]
    seen = {}
    for it in items:
        try:
            price = float(it["price_usd"])
        except (KeyError, TypeError, ValueError):
            rep.warnings.append("%s: skipped an entry with no usable price_usd" % src["id"])
            continue
        blob = norm(it.get("gpu_model")) + " " + norm(it.get("sku_or_instance"))
        model = next((m for m in ("h100", "h200", "b200") if m in blob), None)
        label = "%s %s" % (it.get("sku_or_instance", "?"), it.get("gpu_model", ""))
        if not model:
            rep.unmatched.append("%s: %s (not an H100, H200 or B200 entry)" % (src["id"], label.strip()))
            continue
        sku = norm(it.get("sku_or_instance")) + norm(it.get("gpu_model"))
        cands = [r for r in rows if r["provider"] == src["provider"] and r["gpu"].lower() == model
                 and (not src.get("gpus") or r["gpu"] in src["gpus"]) and any(t in sku for t in r["match_sku"])]
        if len(cands) != 1:
            (rep.warnings if cands else rep.unmatched).append(
                "%s: %s %s" % (src["id"], label.strip(), "matches %d rows, skipped" % len(cands) if cands else "has no matching row"))
            continue
        row = cands[0]
        unit = str(it.get("billing_unit", "")).lower()
        per_instance = bool(re.search(r"instance|node|machine|server|\bvm\b", unit))
        gpus = it.get("gpus_per_instance") or row.get("gpus_per_instance")
        if per_instance:
            if not gpus:
                rep.warnings.append("%s: %s is per instance but no GPU count is known, skipped" % (src["id"], row["label"]))
                continue
            ip, ph = price, price / float(gpus)
        else:
            ph = price
            ip = price * row["gpus_per_instance"] if row.get("gpus_per_instance") else None
        ph = round(ph, 2)
        lo, hi = PLAUSIBLE_PER_GPU_HOUR
        if not lo <= ph <= hi:
            rep.warnings.append("%s: %s extracted %.2f per GPU-hour, outside %.1f to %.1f, skipped" % (src["id"], row["label"], ph, lo, hi))
            continue
        if row["id"] in seen and abs(seen[row["id"]] - ph) > 0.005:
            rep.warnings.append("%s: %s extracted twice with different values, skipped" % (src["id"], row["label"]))
            continue
        seen[row["id"]] = ph
        old = row["price_per_gpu_hour_usd"]
        prev_checked = row.get("checked")
        row["checked"] = today
        if rel_diff(old, ph) <= THRESHOLD:
            rep.unchanged += 1
            continue
        if rel_diff(old, ph) > max_change:
            rep.held.append("%s: %s moved from $%.2f to $%.2f (%.0f percent), above the %.0f percent limit, not applied" % (
                src["id"], row["label"], old, ph, 100 * rel_diff(old, ph), 100 * max_change))
            continue
        row["price_per_gpu_hour_usd"] = ph
        if row.get("instance_price_usd") is not None and ip is not None:
            row["instance_price_usd"] = round(ip, 2)
        row["conditions"] = strip_auto_note(row["conditions"]) + " Previously $%.2f before the %s check." % (old, today)
        if row["gpu"] in HISTORY_MODELS:
            add_history_point(data, row, old, ph, src, today, prev_checked)
        rep.changes.append("%s moved from $%.2f to $%.2f per GPU-hour%s." % (
            row["label"] if row["provider"].split()[-1].lower() in row["label"].lower() else row["provider"] + " " + row["label"], old, ph,
            " (secondary source)" if src["source_type"] == "secondary" else ""))


def apply_storage(data, src, items, today, rep, max_change):
    rows = data["compute"]["storage"]
    for it in items:
        try:
            price = float(it["price_per_gb_month"])
        except (KeyError, TypeError, ValueError):
            rep.warnings.append("%s: skipped a storage entry with no usable price" % src["id"])
            continue
        prod = norm(it.get("product"))
        cands = [r for r in rows if r["provider"] == src["provider"] and any(t in prod for t in r["match_sku"])]
        if len(cands) != 1:
            (rep.warnings if cands else rep.unmatched).append(
                "%s: storage %s %s" % (src["id"], it.get("product", "?"), "matches %d rows, skipped" % len(cands) if cands else "has no matching row"))
            continue
        row = cands[0]
        unit = str(it.get("unit", "")).lower()
        raw = price
        if "hour" in unit:
            price *= HOURS_PER_MONTH
        elif "month" not in unit:
            rep.warnings.append("%s: storage %s has unit '%s', skipped" % (src["id"], row["product"], it.get("unit")))
            continue
        if re.search(r"\btb\b|terabyte", unit):
            rep.warnings.append("%s: storage %s is quoted per TB, skipped" % (src["id"], row["product"]))
            continue
        price = round(price, 4)
        old = row["price_per_month_usd"]
        row["checked"] = today
        if rel_diff(old, price) <= THRESHOLD:
            rep.unchanged += 1
            continue
        if rel_diff(old, price) > max_change:
            rep.held.append("%s: storage %s moved from $%.4f to $%.4f, above the limit, not applied" % (src["id"], row["product"], old, price))
            continue
        row["price_per_month_usd"] = price
        cond = strip_auto_note(row["conditions"])
        if "hour" in unit:
            cond = re.sub(r"Page lists \$[0-9.]+ per GiB-hour", "Page lists $%s per GiB-hour" % ("%.9f" % raw).rstrip("0"), cond)
        row["conditions"] = cond + " Previously $%.4f before the %s check." % (old, today)
        rep.changes.append("%s %s moved from $%.4f to $%.4f per month." % (row["provider"], row["product"], old, price))


# ---------------------------------------------------------------- io
def load_fixture(path, src):
    p = os.path.join(path, src["id"] + ".json") if os.path.isdir(path) else path
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as fh:
        body = json.load(fh)
    if os.path.isfile(path) and src["id"] in body and "success" not in body:
        return body[src["id"]]
    return body


def dump(data):
    return json.dumps(data, indent=1, ensure_ascii=False) + "\n"


def inline_block(data):
    return '<script type="application/json" id="data-inline">%s</script>' % json.dumps(
        data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")


def sync_inline(index_path, data, dry):
    if not os.path.exists(index_path):
        log("index.html not found at %s, inline copy not refreshed" % index_path)
        return
    with open(index_path, encoding="utf-8") as fh:
        html = fh.read()
    if not INLINE_RE.search(html):
        log("No DATA-INLINE markers in index.html, inline copy not refreshed")
        return
    new = INLINE_RE.sub(lambda m: m.group(1) + inline_block(data) + m.group(3), html, count=1)
    if new == html:
        log("Inline copy in index.html already matches data.json")
        return
    if dry:
        log("Dry run, would refresh the inline copy in index.html")
        return
    shutil.copy2(index_path, index_path + ".bak")
    with open(index_path, "w", encoding="utf-8") as fh:
        fh.write(new)
    log("Refreshed the inline copy in index.html (backup index.html.bak)")


# ---------------------------------------------------------------- main
def main(argv=None):
    ap = argparse.ArgumentParser(description="Refresh compute prices in data.json with Firecrawl.")
    ap.add_argument("--data", default=os.path.join(HERE, "data.json"))
    ap.add_argument("--index", default=os.path.join(HERE, "index.html"))
    ap.add_argument("--dry-run", action="store_true", help="print the diff and write nothing")
    ap.add_argument("--fixture", metavar="PATH", help="read saved Firecrawl responses from a directory (one <source id>.json each) or one JSON file keyed by source id, and make no network calls")
    ap.add_argument("--today", help="override today's date, YYYY-MM-DD (useful with --fixture)")
    ap.add_argument("--only", nargs="*", metavar="ID", help="check only these price source ids")
    ap.add_argument("--max-change", type=float, default=0.5, help="hold back any single move larger than this fraction for human review (default 0.5)")
    ap.add_argument("--no-inline", action="store_true", help="do not refresh the inline copy in index.html")
    ap.add_argument("--sync-only", action="store_true", help="only re-embed data.json into index.html, no scraping, no key needed")
    args = ap.parse_args(argv)

    with open(args.data, encoding="utf-8") as fh:
        data = json.load(fh)

    if args.sync_only:
        sync_inline(args.index, data, args.dry_run)
        return 0

    api_key = os.environ.get("FIRECRAWL_API_KEY", "").strip()
    if not args.fixture and not api_key:
        sys.stderr.write("FIRECRAWL_API_KEY is not set. Export your Firecrawl API key and run again, "
                         "or use --fixture to test against saved responses.\n")
        return 2

    today = args.today or datetime.date.today().isoformat()
    original = copy.deepcopy(data)
    before_text = dump(original)
    rep = Report()
    ok_sources = 0

    for src in data["price_sources"]:
        if args.only and src["id"] not in args.only:
            continue
        try:
            resp = load_fixture(args.fixture, src) if args.fixture else scrape(src, today, api_key)
            if resp is None:
                log("skip   %-20s no fixture file" % src["id"])
                continue
            js = extract(resp)
        except Exception as exc:                       # keep going, report at the end
            rep.errors.append("%s: %s" % (src["id"], exc))
            src["last_status"] = "error"
            log("error  %-20s %s" % (src["id"], exc))
            continue
        if "gpu" in src["kinds"]:
            apply_gpu(data, src, js.get("gpu_prices") or [], today, rep, args.max_change)
        if "storage" in src["kinds"]:
            apply_storage(data, src, js.get("storage_prices") or [], today, rep, args.max_change)
        src["last_checked"] = today
        src["last_status"] = "ok"
        ok_sources += 1
        log("ok     %-20s %s" % (src["id"], src["url"]))

    if ok_sources:
        data["compute"]["asof"] = today
        data["last_price_check"] = today
    if rep.changes:
        data["changelog"].append({"date": today, "title": "Price check", "items": list(rep.changes)})

    # guard rails, editorial sections must be untouched and no row may disappear
    for key in PROTECTED:
        if data.get(key) != original.get(key):
            sys.stderr.write("Refusing to continue, the editorial section '%s' would change.\n" % key)
            return 3
    for key in ("gpu", "storage"):
        if len(data["compute"][key]) < len(original["compute"][key]):
            sys.stderr.write("Refusing to continue, a %s row would be removed.\n" % key)
            return 3

    after_text = dump(data)
    log()
    if rep.changes:
        log("Changes applied")
        for c in rep.changes:
            log("  " + c)
    else:
        log("No matched price moved by more than %.1f percent (%d rows checked and unchanged)." % (100 * THRESHOLD, rep.unchanged))
    for title, items in (("Held for review", rep.held), ("Unmatched entries, no row changed", rep.unmatched), ("Warnings", rep.warnings), ("Errors", rep.errors)):
        if items:
            log("\n%s" % title)
            for i in items:
                log("  " + i)

    diff = list(difflib.unified_diff(before_text.splitlines(), after_text.splitlines(), "data.json (current)", "data.json (updated)", n=1, lineterm=""))
    log("\nDiff")
    log("\n".join(diff) if diff else "  no differences")

    if args.dry_run:
        log("\nDry run, nothing written.")
    elif after_text != before_text:
        shutil.copy2(args.data, args.data + ".bak")
        with open(args.data, "w", encoding="utf-8") as fh:
            fh.write(after_text)
        log("\nWrote %s (backup %s.bak)" % (args.data, os.path.basename(args.data)))
    else:
        log("\ndata.json already up to date, nothing written.")

    if not args.no_inline and (not args.dry_run or after_text != before_text):
        sync_inline(args.index, data, args.dry_run)
    return 1 if rep.errors else 0


if __name__ == "__main__":
    sys.exit(main())
