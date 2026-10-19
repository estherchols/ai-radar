"""
Pricing switch model for an AI coding agent (a Cursor-shaped vendor).
Esther Suebsubanant, October 2026. MIT licence. Run: python3 model.py

Simulates 10,000 individual developers for 36 months under four pricing
regimes and reports revenue per user, gross margin, churn, bill
predictability and vendor revenue volatility. Every input is an
assumption and is listed at the top so you can argue with it.

Outputs: model.json (all results), model.csv (headline table),
chart-regimes.svg and chart-fee.svg (inline charts for the web page).
"""
import json, csv, math
import numpy as np

rng = np.random.default_rng(7)
N = 10_000
MONTHS = 36

# ----------------------------------------------------------------- assumptions
# Three developer segments. usage = monthly agent usage valued at model API
# list price, prs = merged agent-authored pull requests a month, wtp = what the
# developer will pay a month before looking for an alternative.
SEG = {
    "light": dict(share=0.45, usage=8.0,   prs=4,  wtp=34.0),
    "daily": dict(share=0.40, usage=35.0,  prs=18, wtp=70.0),
    "heavy": dict(share=0.15, usage=150.0, prs=50, wtp=240.0),
}
USER_SIGMA = 0.35        # persistent spread of usage and willingness to pay within a segment (lognormal)
MONTH_SIGMA = 0.45       # month-to-month spread of an individual's usage (lognormal)
TP_SHARE_FREE = 0.55     # share of usage on third-party models (Claude, GPT, Gemini) when model choice is not penalised
TP_SHARE_BUNDLE = 0.30   # the same share once third-party usage is billed separately (bundle regime)
FP_COST_RATIO = 0.35     # vendor's cost to serve first-party model usage, as a share of list price
TP_COST_RATIO = 1.00     # third-party usage is paid to the provider at list
OTHER_COGS = 4.0         # hosting, support, payments, per active user per month ($)
COST_SHOCK_SIGMA = 0.20  # month-to-month lognormal shock to model list prices (frontier price moves)
CHURN_BASE = 0.020       # background monthly churn
CHURN_OVERPAY = 0.15     # extra monthly churn per 100% that the bill exceeds willingness to pay
CHURN_SHOCK = 0.04       # extra monthly churn when the bill jumps more than 30% month over month
PR_STRICTNESS = 0.85     # share of merged PRs that pass the "agent-authored, merged, lightly edited" test

REGIMES = {
    # Flat seat: $40 a month, any model, fair-use throttle at $200 of list usage.
    "Flat seat":         dict(fee=40.0, pool=200.0, markup=1.00, floor=0.0,  pr_fee=0.0, pr_cap=0.0),
    # Bundle (Cursor today, Premium tier): $30, a $40 pool of first-party usage, first-party overage
    # at list, third-party at list from the first request.
    "Bundle (today)":    dict(fee=30.0, pool=40.0,  markup=1.00, floor=0.0,  pr_fee=0.0, pr_cap=0.0),
    # Pure usage: no fee, every dollar of list usage billed at a markup (calibrated below).
    "Pure usage":        dict(fee=0.0,  pool=0.0,   markup=1.30, floor=0.0,  pr_fee=0.0, pr_cap=0.0),
    # Three-part tariff: platform fee, a compute floor of any-model usage included, usage above the
    # floor at a thin markup, and a success fee per merged agent-authored PR (calibrated below), capped.
    "Three-part tariff": dict(fee=15.0, pool=0.0,   markup=1.10, floor=10.0, pr_fee=2.50, pr_cap=150.0),
}

# ----------------------------------------------------------------- population
names = list(SEG)
idx = rng.choice(len(names), size=N, p=[SEG[s]["share"] for s in names])
def seg(key): return np.array([SEG[s][key] for s in names])[idx]
mp = lambda s: -0.5 * s * s                                  # mean-preserving lognormal offset
base_usage = seg("usage") * np.exp(rng.normal(mp(USER_SIGMA), USER_SIGMA, N))
wtp        = seg("wtp")   * np.exp(rng.normal(mp(USER_SIGMA), USER_SIGMA, N))
pr_rate    = seg("prs")

def simulate(name, R, tp_share, pr_strictness=PR_STRICTNESS, seed=11):
    r = np.random.default_rng(seed)
    alive = np.ones(N, dtype=bool)
    prev_bill = np.full(N, np.nan)
    rev, cogs, churn, alive_n = [], [], [], []
    bills = np.full((MONTHS, N), np.nan)
    for m in range(MONTHS):
        shock = math.exp(r.normal(mp(COST_SHOCK_SIGMA), COST_SHOCK_SIGMA))
        usage = base_usage * np.exp(r.normal(mp(MONTH_SIGMA), MONTH_SIGMA, N)) * shock
        if name == "Flat seat":
            usage = np.minimum(usage, R["pool"])
        tp = usage * r.beta(2.0, 2.0 / tp_share - 2.0, N)
        fp = usage - tp
        prs = r.poisson(pr_rate)
        counted = r.binomial(prs, pr_strictness)

        if name == "Flat seat":
            bill = np.full(N, R["fee"])
        elif name == "Bundle (today)":
            bill = R["fee"] + np.maximum(0.0, fp - R["pool"]) * R["markup"] + tp * R["markup"]
        elif name == "Pure usage":
            bill = usage * R["markup"]
        else:
            bill = R["fee"] + np.maximum(0.0, usage - R["floor"]) * R["markup"] + np.minimum(counted * R["pr_fee"], R["pr_cap"])

        cost = fp * FP_COST_RATIO + tp * TP_COST_RATIO + OTHER_COGS
        over = np.maximum(0.0, bill / wtp - 1.0)
        jump = np.where(np.isnan(prev_bill), 0.0, np.maximum(0.0, bill / np.maximum(prev_bill, 1e-9) - 1.30))
        p = CHURN_BASE + CHURN_OVERPAY * over + CHURN_SHOCK * (jump > 0)
        churned = alive & (r.random(N) < np.minimum(p, 0.9))

        rev.append(bill[alive].sum()); cogs.append(cost[alive].sum()); alive_n.append(alive.sum())
        churn.append(churned.sum() / max(alive.sum(), 1))
        bills[m, alive] = bill[alive]
        prev_bill = bill
        alive &= ~churned

    rev, cogs, alive_n = map(np.array, (rev, cogs, alive_n))
    w = slice(6, MONTHS)
    full = ~np.isnan(bills[w]).any(axis=0)
    bill_cv = float(np.mean(bills[w][:, full].std(axis=0) / bills[w][:, full].mean(axis=0))) if full.any() else 0.0
    return dict(
        regime=name,
        arpu_m1=float(rev[0] / alive_n[0]),
        arpu=float((rev[w] / alive_n[w]).mean()),
        gross_margin=float(1.0 - cogs[w].sum() / rev[w].sum()),
        monthly_churn=float(np.mean(churn[6:])),
        retained_36m=float(alive_n[-1] / N),
        revenue_36m=float(rev.sum()),
        gross_profit_36m=float((rev - cogs).sum()),
        revenue_cv=float(rev[w].std() / rev[w].mean()),
        bill_cv=bill_cv,
    )

def month1_arpu(name, R, tp_share):
    return simulate(name, R, tp_share)["arpu_m1"]

# Calibrate pure usage markup and the success fee so every regime starts at the bundle's month-one ARPU.
# That isolates the effect of the structure from the effect of simply charging more.
target = month1_arpu("Bundle (today)", REGIMES["Bundle (today)"], TP_SHARE_BUNDLE)
def bisect(fn, lo, hi, it=28):
    for _ in range(it):
        mid = (lo + hi) / 2
        if fn(mid) < target: lo = mid
        else: hi = mid
    return (lo + hi) / 2
R = REGIMES["Pure usage"]
R["markup"] = bisect(lambda k: month1_arpu("Pure usage", dict(R, markup=k), TP_SHARE_FREE), 0.3, 3.0)
R = REGIMES["Three-part tariff"]
R["pr_fee"] = bisect(lambda f: month1_arpu("Three-part tariff", dict(R, pr_fee=f), TP_SHARE_FREE), 0.0, 6.0)

results = [simulate(n, R, TP_SHARE_BUNDLE if n == "Bundle (today)" else TP_SHARE_FREE) for n, R in REGIMES.items()]

T = REGIMES["Three-part tariff"]
sens_fee = [dict(pr_fee=f, **simulate("Three-part tariff", dict(T, pr_fee=f), TP_SHARE_FREE)) for f in (0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0)]
sens_strict = [dict(strictness=s, **simulate("Three-part tariff", T, TP_SHARE_FREE, pr_strictness=s)) for s in (0.5, 0.65, 0.85, 1.0)]
sens_shock = []
for sig in (0.0, 0.2, 0.4):
    COST_SHOCK_SIGMA = sig
    sens_shock.append(dict(shock=sig, rows=[simulate(n, R, TP_SHARE_BUNDLE if n == "Bundle (today)" else TP_SHARE_FREE) for n, R in REGIMES.items()]))
COST_SHOCK_SIGMA = 0.20

out = dict(results=results, calibrated=dict(target_month1_arpu=target, usage_markup=REGIMES["Pure usage"]["markup"], pr_fee=T["pr_fee"]),
           sensitivity_fee=sens_fee, sensitivity_strictness=sens_strict, sensitivity_shock=sens_shock,
           assumptions=dict(N=N, months=MONTHS, segments=SEG, user_sigma=USER_SIGMA, month_sigma=MONTH_SIGMA,
                            tp_share_free=TP_SHARE_FREE, tp_share_bundle=TP_SHARE_BUNDLE, fp_cost_ratio=FP_COST_RATIO,
                            tp_cost_ratio=TP_COST_RATIO, other_cogs=OTHER_COGS, cost_shock_sigma=0.20, churn_base=CHURN_BASE,
                            churn_overpay=CHURN_OVERPAY, churn_shock=CHURN_SHOCK, pr_strictness=PR_STRICTNESS, regimes=REGIMES))
json.dump(out, open("model.json", "w"), indent=1)
with open("model.csv", "w", newline="") as f:
    wr = csv.DictWriter(f, fieldnames=list(results[0])); wr.writeheader(); wr.writerows(results)

# ----------------------------------------------------------------- charts (inline SVG, theme-aware via CSS variables)
def svg_regimes(rows, path):
    W, H, L, B, T0 = 720, 300, 54, 44, 24
    metrics = [("gross_margin", "Gross margin", 100), ("monthly_churn", "Monthly churn", 100), ("revenue_cv", "Revenue volatility", 100), ("bill_cv", "Bill volatility", 100)]
    cols = ["var(--muted)", "var(--gold)", "var(--accent)", "var(--pine)"]
    gw = (W - L - 20) / len(metrics); bw = gw / (len(rows) + 1.2)
    ymax = 70
    s = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Gross margin, churn and volatility by pricing regime" font-family="Inter,system-ui,sans-serif" font-size="11">']
    for y in range(0, ymax + 1, 10):
        yy = T0 + (H - T0 - B) * (1 - y / ymax)
        s.append(f'<line x1="{L}" x2="{W-20}" y1="{yy:.1f}" y2="{yy:.1f}" stroke="var(--line)" stroke-width="1"/><text x="{L-8}" y="{yy+4:.1f}" text-anchor="end" fill="var(--muted)">{y}%</text>')
    for mi, (k, label, scale) in enumerate(metrics):
        x0 = L + mi * gw + bw * 0.6
        for ri, r in enumerate(rows):
            v = r[k] * scale; hgt = (H - T0 - B) * min(v, ymax) / ymax
            x = x0 + ri * bw; y = H - B - hgt
            s.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw*0.86:.1f}" height="{hgt:.1f}" rx="3" fill="{cols[ri]}"><title>{r["regime"]}: {label} {v:.1f}%</title></rect>')
            s.append(f'<text x="{x+bw*0.43:.1f}" y="{y-4:.1f}" text-anchor="middle" fill="var(--ink-2)" font-size="10">{v:.0f}</text>')
        s.append(f'<text x="{L + mi*gw + gw/2:.1f}" y="{H-B+18}" text-anchor="middle" fill="var(--ink)" font-weight="500">{label}</text>')
    lx = L
    for ri, r in enumerate(rows):
        s.append(f'<rect x="{lx}" y="{H-14}" width="10" height="10" rx="2" fill="{cols[ri]}"/><text x="{lx+14}" y="{H-5}" fill="var(--ink-2)">{r["regime"]}</text>')
        lx += 16 + 7 * len(r["regime"]) + 14
    s.append('</svg>'); open(path, "w").write("\n".join(s))

def svg_fee(rows, path):
    W, H, L, B, T0 = 720, 280, 54, 40, 20
    fees = [r["pr_fee"] for r in rows]
    s = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Gross profit and retention versus success fee" font-family="Inter,system-ui,sans-serif" font-size="11">']
    gp = [r["gross_profit_36m"] / 1e6 for r in rows]; ret = [r["retained_36m"] * 100 for r in rows]
    gmax = max(gp) * 1.15; xs = lambda f: L + (W - L - 60) * (f - fees[0]) / (fees[-1] - fees[0])
    yg = lambda v: T0 + (H - T0 - B) * (1 - v / gmax); yr = lambda v: T0 + (H - T0 - B) * (1 - v / 100)
    for i in range(0, 5):
        yy = T0 + (H - T0 - B) * i / 4
        s.append(f'<line x1="{L}" x2="{W-60}" y1="{yy:.1f}" y2="{yy:.1f}" stroke="var(--line)"/>')
        s.append(f'<text x="{L-8}" y="{yy+4:.1f}" text-anchor="end" fill="var(--muted)">${gmax*(1-i/4):.0f}M</text>')
        s.append(f'<text x="{W-52}" y="{yy+4:.1f}" fill="var(--muted)">{100*(1-i/4):.0f}%</text>')
    s.append('<polyline fill="none" stroke="var(--pine)" stroke-width="2.5" points="' + " ".join(f"{xs(f):.1f},{yg(v):.1f}" for f, v in zip(fees, gp)) + '"/>')
    s.append('<polyline fill="none" stroke="var(--accent)" stroke-width="2.5" stroke-dasharray="5 4" points="' + " ".join(f"{xs(f):.1f},{yr(v):.1f}" for f, v in zip(fees, ret)) + '"/>')
    for f, v, rv in zip(fees, gp, ret):
        s.append(f'<circle cx="{xs(f):.1f}" cy="{yg(v):.1f}" r="3.5" fill="var(--pine)"><title>${f:.2f} per PR: gross profit ${v:.1f}M over 36 months</title></circle>')
        s.append(f'<circle cx="{xs(f):.1f}" cy="{yr(rv):.1f}" r="3.5" fill="var(--accent)"><title>${f:.2f} per PR: {rv:.0f}% of developers retained after 36 months</title></circle>')
        s.append(f'<text x="{xs(f):.1f}" y="{H-B+16}" text-anchor="middle" fill="var(--ink-2)">${f:g}</text>')
    s.append(f'<text x="{(L+W-60)/2:.0f}" y="{H-6}" text-anchor="middle" fill="var(--ink)" font-weight="500">Success fee per merged agent-authored PR</text>')
    s.append(f'<text x="{L}" y="{T0-6}" fill="var(--pine)" font-weight="500">— Gross profit, 36 months (left axis)</text><text x="{L+230}" y="{T0-6}" fill="var(--accent)" font-weight="500">- - Developers retained at month 36 (right axis)</text>')
    s.append('</svg>'); open(path, "w").write("\n".join(s))

svg_regimes(results, "chart-regimes.svg")
svg_fee(sens_fee, "chart-fee.svg")

print(f"calibrated: month-1 ARPU target ${target:.2f}; usage markup {REGIMES['Pure usage']['markup']:.2f}x; success fee ${T['pr_fee']:.2f}/PR")
for r in results:
    print(f'{r["regime"]:18s} ARPU ${r["arpu"]:6.2f}  GM {r["gross_margin"]*100:5.1f}%  churn/mo {r["monthly_churn"]*100:4.1f}%  '
          f'retained {r["retained_36m"]*100:4.1f}%  rev36 ${r["revenue_36m"]/1e6:5.2f}M  GP36 ${r["gross_profit_36m"]/1e6:5.2f}M  revCV {r["revenue_cv"]:.3f}  billCV {r["bill_cv"]:.2f}')
print("fee:", [(s["pr_fee"], round(s["arpu"],1), round(s["gross_margin"],3), round(s["retained_36m"],3), round(s["gross_profit_36m"]/1e6,2)) for s in sens_fee])
print("strict:", [(s["strictness"], round(s["arpu"],1), round(s["gross_margin"],3), round(s["retained_36m"],3)) for s in sens_strict])
for blk in sens_shock:
    print("shock", blk["shock"], [(r["regime"], round(r["gross_margin"],3), round(r["revenue_cv"],3), round(r["retained_36m"],3)) for r in blk["rows"]])
