"""Inline the SVG charts and headline numbers from model.json into index.html.
Run after model.py:  python3 build.py
Placeholders in index.html:  <!--CHART:regimes-->  <!--CHART:fee-->  and {{key}} tokens listed in TOKENS.
The source template is index.template.html; index.html is generated."""
import json
d = json.load(open("model.json"))
rows = {r["regime"]: r for r in d["results"]}
pct = lambda v: f"{v*100:.0f}%"
pct1 = lambda v: f"{v*100:.1f}%"
usd = lambda v: f"${v:,.0f}"
T = {}
for key, r in rows.items():
    k = {"Flat seat": "seat", "Bundle (today)": "bundle", "Pure usage": "usage", "Three-part tariff": "tariff"}[key]
    T[f"{k}_arpu"] = f"${r['arpu']:.0f}"; T[f"{k}_gm"] = pct(r["gross_margin"]); T[f"{k}_churn"] = pct1(r["monthly_churn"])
    T[f"{k}_ret"] = pct(r["retained_36m"]); T[f"{k}_rev"] = f"${r['revenue_36m']/1e6:.1f}M"; T[f"{k}_gp"] = f"${r['gross_profit_36m']/1e6:.1f}M"
    T[f"{k}_revcv"] = pct(r["revenue_cv"]); T[f"{k}_billcv"] = pct(r["bill_cv"])
T["target_arpu"] = f"${d['calibrated']['target_month1_arpu']:.0f}"
T["usage_markup"] = f"{d['calibrated']['usage_markup']:.2f}x"
T["pr_fee"] = f"${d['calibrated']['pr_fee']:.2f}"
best = max(d["sensitivity_fee"], key=lambda s: s["gross_profit_36m"])
T["best_fee"] = f"${best['pr_fee']:.2f}"; T["best_fee_gp"] = f"${best['gross_profit_36m']/1e6:.1f}M"; T["best_fee_ret"] = pct(best["retained_36m"])
hi = d["sensitivity_fee"][-1]; T["hi_fee"] = f"${hi['pr_fee']:.2f}"; T["hi_fee_ret"] = pct(hi["retained_36m"]); T["hi_fee_gp"] = f"${hi['gross_profit_36m']/1e6:.1f}M"
s0, s1 = d["sensitivity_strictness"][0], d["sensitivity_strictness"][-1]
T["strict_lo_gm"] = pct(s0["gross_margin"]); T["strict_hi_gm"] = pct(s1["gross_margin"])
sh = {b["shock"]: {r["regime"]: r for r in b["rows"]} for b in d["sensitivity_shock"]}
T["shock_bundle_cv"] = pct(sh[0.4]["Bundle (today)"]["revenue_cv"]); T["shock_tariff_cv"] = pct(sh[0.4]["Three-part tariff"]["revenue_cv"]); T["shock_usage_cv"] = pct(sh[0.4]["Pure usage"]["revenue_cv"])

html = open("index.template.html").read()
html = html.replace("<!--CHART:regimes-->", open("chart-regimes.svg").read()).replace("<!--CHART:fee-->", open("chart-fee.svg").read())
for k, v in T.items():
    html = html.replace("{{" + k + "}}", v)
missing = [t for t in set(__import__("re").findall(r"\{\{(\w+)\}\}", html))]
assert not missing, f"unfilled tokens: {missing}"
open("index.html", "w").write(html)
print("built index.html with", len(T), "tokens")
