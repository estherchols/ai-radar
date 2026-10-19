"""Render research/fs-gpu/index.html from the recalculated model.

Every number in the article that comes from the model is read from fs-gpu-model.xlsx
(after recalc), so the page and the workbook cannot drift apart. Run build_model.py,
then recalc.py, then this.
"""
import math
from openpyxl import load_workbook

HERE = "/home/claude/ai-radar/research/fs-gpu/"
wb = load_workbook(HERE + "fs-gpu-model.xlsx", data_only=True)

def by_code(sheet, code, col=3):
    ws = wb[sheet]
    for r in ws.iter_rows(min_row=5):
        if r[0].value == code:
            return r[col - 1].value
    raise KeyError(f"{sheet}:{code}")

def by_label(sheet, label, col=3):
    ws = wb[sheet]
    for r in ws.iter_rows(min_row=5):
        if r[0].value == label or r[1].value == label:
            return r[col - 1].value
    raise KeyError(f"{sheet}:{label}")

V = {}
# TopDown
for k in ["A3", "A6", "A8", "A10", "A12", "B3", "B5", "B6", "C3", "C5", "C7", "R4", "R5", "R6", "R7", "R7b", "R7d", "R7e", "R8"]:
    V["td_" + k] = by_code("TopDown", k)
# BottomUp rows
bu = wb["BottomUp"]
segs = []
for r in bu.iter_rows(min_row=5, max_row=12):
    segs.append(dict(name=r[0].value, n=r[1].value, adopt=r[2].value, own=r[3].value, adopters=r[4].value,
                     per=r[5].value, gpus=r[6].value, price=r[7].value, value=r[8].value, repl=r[9].value, grow=r[10].value, spend=r[11].value))
tot = [c.value for c in bu[13]]
V["bu_gpus"] = tot[6]; V["bu_value"] = tot[8]; V["bu_spend"] = tot[11]; V["bu_adopters"] = tot[4]
V["bu_rent"] = bu["C31"].value; V["bu_total_b"] = bu["C32"].value; V["bu_nvda_b"] = bu["C33"].value; V["bu_quantpay"] = bu["C34"].value
proj = []
for r in bu.iter_rows(min_row=38, max_row=42):
    if r[0].value and str(r[0].value).isdigit():
        proj.append((r[0].value, r[1].value, r[2].value, r[5].value))
V["proj"] = proj
# Workloads
for k in ["W4", "W5", "W6", "W8", "W9", "W10", "W11", "W12", "F2", "F5", "F7", "F10", "F11", "Q2", "Q3"]:
    V["wl_" + k] = by_code("Workloads", k)
# TCO
for k in ["T3", "T4", "T5", "T6", "T7", "T8", "T9", "T10", "T11", "T12", "T13", "T14", "T15", "T16", "T17", "T18", "U1", "U2", "U3", "U4", "U5", "U6", "U7", "U8", "S1", "S2", "S3", "S4", "S5"]:
    V["tco_" + k] = by_code("TCO", k)
sens = []
for r in wb["TCO"].iter_rows(min_row=5):
    if r[1].value and str(r[1].value).startswith("Utilisation "):
        sens.append((r[1].value.replace("Utilisation ", ""), r[2].value))
V["sens"] = sens
# Value
for k in ["VA2", "VA3", "VA4", "VA6", "VA7", "VA8", "VA9", "VA10", "VB3", "VB4", "VB5", "VB6", "VB7", "VB9", "VC3", "VC5", "VC6", "VC7", "VC8", "VC9", "VC10", "VD2", "VD3"]:
    V["val_" + k] = by_code("Value", k)
# Configs
cfg_fmt = {"K3": lambda x: f"{x:,.0f}", "K4": lambda x: f"{x:,.0f}", "K5": lambda x: f"{x:,.0f}", "K9": lambda x: f"{x:,.0f}", "K14": lambda x: f"${x/1e6:,.1f}M", "K19": lambda x: f"${x/1e6:,.2f}M", "K21": lambda x: f"${x:,.1f}M", "K22": lambda x: f"{x*100:.0f}%", "K24": lambda x: f"${x:,.2f}", "K26": lambda x: f"${x:,.0f}", "K27": lambda x: f"{x*100:.0f}%", "K29": lambda x: f"{x*100:.0f}%", "K30": lambda x: f"{x:,.1f}", "K31": lambda x: f"{x:,.1f}", "K35": lambda x: f"${x:,.0f}M", "K37": lambda x: f"{x:,.0f}x"}
CF = {}
for code, f in cfg_fmt.items():
    for i, col in enumerate((3, 4, 5, 6), start=1):
        v = by_code("Configs", code, col)
        CF[f"cf_{code}_{i}"] = f(v) if isinstance(v, (int, float)) else str(v)
CF["cf_K35_3"] = f"${by_code('Configs','K35',5)/1000:,.1f}B"
# Pricing
for k in ["P1", "P2", "P7", "M1", "M2", "M3", "M4", "M8", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "H3", "H4"]:
    V["pr_" + k] = by_code("Pricing", k)

# ------------------------------------------------------------- formatters
def usd(x, d=0):
    return f"${x:,.{d}f}"
def bn(x, d=1):
    return f"${x:,.{d}f}B"
def mn(x, d=0):
    return f"${x:,.{d}f}M"
def pct(x, d=0):
    return f"{x*100:,.{d}f}%"
def num(x, d=0):
    return f"{x:,.{d}f}"
def k(x):
    return f"{x/1000:,.0f},000" if x >= 1000 else f"{x:,.0f}"

# ------------------------------------------------------------- SVG charts
INK = "var(--ink)"; DIM = "var(--dim)"; LINE = "var(--line)"; ACC = "var(--electric)"; ACC2 = "#e3a248"; ACC3 = "#c4561c"

def chart_sizing():
    bars = [("Path A\nIT budget", V["td_A10"]), ("Path B\nAI cloud", V["td_B6"]), ("Path C\nAI budget", V["td_C7"]),
            ("Top-down\naverage", V["td_R4"]), ("Top-down,\nGPU only", V["td_R7d"]), ("Bottom-up\nby segment", V["td_R6"])]
    W, H = 720, 300; L, R, T, B = 50, 20, 20, 60
    mx = 30
    pw = W - L - R; ph = H - T - B
    bw = pw / len(bars) * 0.6
    out = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Top-down versus bottom-up sizing of financial services GPU spend 2026">']
    for g in range(0, 31, 10):
        y = T + ph - ph * g / mx
        out.append(f'<line x1="{L}" y1="{y:.1f}" x2="{W-R}" y2="{y:.1f}" stroke="{LINE}" stroke-width="1"/>')
        out.append(f'<text x="{L-8}" y="{y+4:.1f}" font-size="11" text-anchor="end" fill="{DIM}" font-family="IBM Plex Mono,monospace">${g}B</text>')
    for i, (lab, val) in enumerate(bars):
        x = L + pw / len(bars) * (i + 0.2)
        h = ph * val / mx
        y = T + ph - h
        fill = ACC if i < 3 else (ACC2 if i in (3, 4) else ACC3)
        out.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{h:.1f}" rx="3" fill="{fill}" opacity="{0.9 if i<3 else 1}"/>')
        out.append(f'<text x="{x+bw/2:.1f}" y="{y-6:.1f}" font-size="12" text-anchor="middle" fill="{INK}" font-weight="600" font-family="Schibsted Grotesk,sans-serif">${val:,.1f}B</text>')
        for j, line in enumerate(lab.split("\n")):
            out.append(f'<text x="{x+bw/2:.1f}" y="{T+ph+16+j*13}" font-size="11" text-anchor="middle" fill="{DIM}" font-family="Schibsted Grotesk,sans-serif">{line}</text>')
    out.append("</svg>")
    return "\n".join(out)

def chart_segments():
    rows = sorted(segs, key=lambda s: -s["gpus"])
    W = 720; rowh = 30; L = 250; R = 90; T = 10
    H = T + rowh * len(rows) + 10
    mx = max(s["gpus"] for s in rows)
    out = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Installed GPUs by financial services segment 2026">']
    for i, s in enumerate(rows):
        y = T + i * rowh
        w = (W - L - R) * s["gpus"] / mx
        out.append(f'<text x="{L-10}" y="{y+19}" font-size="12" text-anchor="end" fill="{INK}" font-family="Schibsted Grotesk,sans-serif">{s["name"].replace("(ranks 101-500)","").strip()}</text>')
        out.append(f'<rect x="{L}" y="{y+6}" width="{w:.1f}" height="18" rx="3" fill="{ACC if i>1 else ACC3}"/>')
        out.append(f'<text x="{L+w+8:.1f}" y="{y+19}" font-size="12" fill="{INK}" font-family="IBM Plex Mono,monospace">{s["gpus"]:,.0f}</text>')
    out.append("</svg>")
    return "\n".join(out)

def chart_tco():
    W, H = 720, 300; L, R, T, B = 55, 20, 20, 45
    pw = W - L - R; ph = H - T - B
    us = [x / 100 for x in range(20, 101, 5)]
    tco = V["tco_T12"]; hrs = V["tco_T15"]
    ys = [tco / (hrs * u) for u in us]
    ymax = 10
    def X(u): return L + pw * (u - 0.2) / 0.8
    def Y(v): return T + ph - ph * min(v, ymax) / ymax
    out = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Effective cost per GPU hour of an owned Blackwell node versus utilisation, against cloud prices">']
    for g in range(0, 11, 2):
        out.append(f'<line x1="{L}" y1="{Y(g):.1f}" x2="{W-R}" y2="{Y(g):.1f}" stroke="{LINE}"/>')
        out.append(f'<text x="{L-8}" y="{Y(g)+4:.1f}" font-size="11" text-anchor="end" fill="{DIM}" font-family="IBM Plex Mono,monospace">${g}</text>')
    for u in (0.2, 0.4, 0.6, 0.8, 1.0):
        out.append(f'<text x="{X(u):.1f}" y="{T+ph+18}" font-size="11" text-anchor="middle" fill="{DIM}" font-family="IBM Plex Mono,monospace">{int(u*100)}%</text>')
    out.append(f'<text x="{L+pw/2:.1f}" y="{H-6}" font-size="11" text-anchor="middle" fill="{DIM}" font-family="Schibsted Grotesk,sans-serif">Utilisation of the owned GPUs (share of hours busy)</text>')
    pts = " ".join(f"{X(u):.1f},{Y(y):.1f}" for u, y in zip(us, ys))
    out.append(f'<polyline points="{pts}" fill="none" stroke="{ACC}" stroke-width="2.5"/>')
    od = 6.99; res = 4.75
    out.append(f'<line x1="{L}" y1="{Y(od):.1f}" x2="{W-R}" y2="{Y(od):.1f}" stroke="{ACC3}" stroke-width="1.5" stroke-dasharray="6 4"/>')
    out.append(f'<text x="{W-R-4}" y="{Y(od)-6:.1f}" font-size="11" text-anchor="end" fill="{ACC3}" font-family="Schibsted Grotesk,sans-serif">Cloud on-demand B200 $6.99</text>')
    out.append(f'<line x1="{L}" y1="{Y(res):.1f}" x2="{W-R}" y2="{Y(res):.1f}" stroke="{ACC2}" stroke-width="1.5" stroke-dasharray="6 4"/>')
    out.append(f'<text x="{W-R-4}" y="{Y(res)-6:.1f}" font-size="11" text-anchor="end" fill="{ACC2}" font-family="Schibsted Grotesk,sans-serif">Cloud reserved B200 $4.75</text>')
    be1 = V["tco_U6"]; be2 = V["tco_U7"]
    for be, col in ((be1, ACC3), (be2, ACC2)):
        out.append(f'<circle cx="{X(be):.1f}" cy="{Y(tco/(hrs*be)):.1f}" r="5" fill="{col}"/>')
        out.append(f'<text x="{X(be):.1f}" y="{Y(tco/(hrs*be))+20:.1f}" font-size="11" text-anchor="middle" fill="{col}" font-family="IBM Plex Mono,monospace">{be*100:.0f}%</text>')
    out.append(f'<text x="{X(0.65)+8:.1f}" y="{Y(V["tco_T17"])-10:.1f}" font-size="11" fill="{ACC}" font-family="Schibsted Grotesk,sans-serif">Owned node, five-year TCO</text>')
    out.append("</svg>")
    return "\n".join(out)

def chart_matrix():
    W, H = 720, 520; L, T = 70, 30; pw, ph = 600, 420
    out = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Two by two map of financial services GPU buyers">']
    out.append(f'<rect x="{L}" y="{T}" width="{pw}" height="{ph}" fill="none" stroke="{LINE}"/>')
    out.append(f'<line x1="{L+pw/2}" y1="{T}" x2="{L+pw/2}" y2="{T+ph}" stroke="{LINE}"/>')
    out.append(f'<line x1="{L}" y1="{T+ph/2}" x2="{L+pw}" y2="{T+ph/2}" stroke="{LINE}"/>')
    # axes labels
    out.append(f'<text x="{L+pw/2}" y="{T+ph+28}" font-size="12" text-anchor="middle" fill="{DIM}" font-family="Schibsted Grotesk,sans-serif">Compute intensity per institution (GPUs installed)  →</text>')
    out.append(f'<text transform="translate(22,{T+ph/2}) rotate(-90)" font-size="12" text-anchor="middle" fill="{DIM}" font-family="Schibsted Grotesk,sans-serif">Pull toward owned, in-jurisdiction infrastructure (regulation, data)  →</text>')
    out.append(f'<text x="{L+6}" y="{T+ph+14}" font-size="10" fill="{DIM}" font-family="IBM Plex Mono,monospace">tens to hundreds</text>')
    out.append(f'<text x="{L+pw-6}" y="{T+ph+14}" font-size="10" text-anchor="end" fill="{DIM}" font-family="IBM Plex Mono,monospace">thousands to tens of thousands</text>')
    quads = [
        (L + pw / 2, T, "Sovereign fleets", "G-SIBs · payment networks · large insurers", "Own, in region. Buy SuperPOD or NVL72 through Dell, HPE or a sovereign cloud. Train transaction and risk models, serve open models to staff.", f"~{(segs[0]['gpus']+segs[3]['gpus']+segs[6]['gpus'])/1000:.0f}k GPUs", ACC3),
        (L, T, "Regulated pods", "Regional banks · mid-size insurers · exchanges", "Small owned pods (4 to 8 nodes) for documents, fraud and compliance. NIM blueprints, EDB or Cloudera on the bank's own data.", f"~{(segs[2]['gpus']+segs[7]['gpus'])/1000:.0f}k GPUs", ACC2),
        (L, T + ph / 2, "Renters", "Hedge funds · asset managers · fintechs", "Rent reserved cloud capacity. Price sensitive, elastic, no residency pull. Compete on $/GPU-hour and reserved terms.", f"~{segs[5]['gpus']/1000:.0f}k GPUs", DIM),
        (L + pw / 2, T + ph / 2, "Research factories", "Quant and market-making firms", "Own, but for latency, cost and secrecy rather than regulation. Build their own data centers (XTX Finland, HRT Lefdal). First to Rubin.", f"~{segs[4]['gpus']/1000:.0f}k GPUs", ACC),
    ]
    def wrap(text, n):
        words = text.split(); lines = []; cur = ""
        for w in words:
            if len(cur) + len(w) + 1 > n and cur:
                lines.append(cur); cur = w
            else:
                cur = (cur + " " + w).strip()
        lines.append(cur)
        return lines
    for x, y, title, who, body, gp, col in quads:
        out.append(f'<rect x="{x+10}" y="{y+10}" width="{pw/2-20}" height="{ph/2-20}" rx="10" fill="var(--bg)" stroke="{col}" stroke-width="1.5"/>')
        out.append(f'<text x="{x+24}" y="{y+40}" font-size="17" font-weight="700" fill="{INK}" font-family="Schibsted Grotesk,sans-serif">{title}</text>')
        out.append(f'<text x="{x+pw/2-24}" y="{y+40}" font-size="11.5" text-anchor="end" fill="{col}" font-family="IBM Plex Mono,monospace">{gp}</text>')
        yy = y + 60
        for ln in wrap(who, 36):
            out.append(f'<text x="{x+24}" y="{yy}" font-size="11" fill="{col}" font-family="IBM Plex Mono,monospace">{ln}</text>')
            yy += 15
        yy += 8
        for ln in wrap(body, 38)[:8]:
            out.append(f'<text x="{x+24}" y="{yy}" font-size="12.5" fill="{INK}" font-family="Schibsted Grotesk,sans-serif">{ln}</text>')
            yy += 17
    out.append("</svg>")
    return "\n".join(out)

def chart_value():
    rows = [("Payments fraud: fleet cost vs fraud prevented", V["val_VA7"]),
            ("G-SIB assistant: GPU cost vs reported savings", V["val_VB7"]),
            ("Quant fleet: 25,000 GPUs vs top-firm trading revenue", V["val_VD3"]),
            ("Risk grid: GPU capital vs five-year CPU refresh avoided", V["val_VC9"] / V["val_VC8"])]
    W = 720; rowh = 34; L = 330; R = 80; T = 10
    H = T + rowh * len(rows) + 10
    mx = max(v for _, v in rows)
    out = [f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Annual compute cost as a share of value created by use case">']
    for i, (lab, v) in enumerate(rows):
        y = T + i * rowh
        w = max(2, (W - L - R) * v / mx)
        out.append(f'<text x="{L-10}" y="{y+21}" font-size="12" text-anchor="end" fill="{INK}" font-family="Schibsted Grotesk,sans-serif">{lab}</text>')
        out.append(f'<rect x="{L}" y="{y+8}" width="{w:.1f}" height="18" rx="3" fill="{ACC}"/>')
        out.append(f'<text x="{L+w+8:.1f}" y="{y+21}" font-size="12" fill="{INK}" font-family="IBM Plex Mono,monospace">{v*100:.1f}%</text>')
    out.append("</svg>")
    return "\n".join(out)

for name, fn in [("chart-sizing", chart_sizing), ("chart-segments", chart_segments), ("chart-tco", chart_tco), ("chart-matrix", chart_matrix), ("chart-value", chart_value)]:
    open(HERE + name + ".svg", "w").write(fn())

seg_rows = "\n".join(
    f'<tr><td><b>{s["name"]}</b></td><td class="num">{s["n"]:,.0f}</td><td class="num">{s["adopt"]*100:.0f}%</td><td class="num">{s["own"]*100:.0f}%</td><td class="num">{s["per"]:,.0f}</td><td class="num">{s["gpus"]:,.0f}</td><td class="num">${s["value"]:,.0f}M</td><td class="num">${s["spend"]:,.0f}M</td></tr>'
    for s in segs)
proj_rows = "\n".join(f'<tr><td class="num">{y}</td><td class="num">{g:,.0f}</td><td class="num">${v/1000:,.1f}B</td><td class="num">${s/1000:,.1f}B</td></tr>' for y, g, v, s in V["proj"])
sens_rows = "".join(f'<tr><td class="num">{u}</td><td class="num">${v:,.2f}</td></tr>' for u, v in V["sens"])

tmpl = open(HERE + "index.template.html").read()
rep = {
    "SEG_ROWS": seg_rows, "PROJ_ROWS": proj_rows, "SENS_ROWS": sens_rows,
    "CHART_SIZING": chart_sizing(), "CHART_SEGMENTS": chart_segments(), "CHART_TCO": chart_tco(), "CHART_MATRIX": chart_matrix(), "CHART_VALUE": chart_value(),
    "td_A3": pct(V["td_A3"], 1), "td_A6": bn(V["td_A6"], 0), "td_A8": bn(V["td_A10"], 1), "td_A10": bn(V["td_A10"], 1), "td_A12": bn(V["td_A12"], 1),
    "td_B3": bn(V["td_B3"], 1), "td_B5": bn(V["td_B5"], 1), "td_B6": bn(V["td_B6"], 1),
    "td_C3": bn(V["td_C3"], 1), "td_C5": bn(V["td_C5"], 1), "td_C7": bn(V["td_C7"], 1),
    "td_R4": bn(V["td_R4"], 1), "td_R5": pct(V["td_R5"]), "td_R6": bn(V["td_R6"], 1), "td_R7": pct(V["td_R7"]), "td_R7b": bn(V["td_R7b"], 1), "td_R7d": bn(V["td_R7d"], 1), "td_R7e": pct(V["td_R7e"]), "td_R8": pct(V["td_R8"], 1),
    "bu_gpus": num(V["bu_gpus"]), "bu_gpus_k": f"{V['bu_gpus']/1000:,.0f},000", "bu_value": bn(V["bu_value"] / 1000, 1), "bu_spend": bn(V["bu_spend"] / 1000, 1), "bu_rent": bn(V["bu_rent"] / 1000, 1), "bu_total": bn(V["bu_total_b"], 1), "bu_nvda": bn(V["bu_nvda_b"], 1), "bu_quantpay": pct(V["bu_quantpay"]), "bu_adopters": num(V["bu_adopters"]),
    "proj_2030_gpus": num(V["proj"][-1][1]), "proj_2030_spend": bn(V["proj"][-1][3] / 1000, 1),
    "wl_W4": f"{V['wl_W4']/1e9:,.1f} billion", "wl_W6": num(V["wl_W6"]), "wl_W8": num(V["wl_W8"]), "wl_W9": pct(V["wl_W9"]), "wl_W10": f"{V['wl_W10']/1000:,.1f} trillion", "wl_W11": f"${V['wl_W11']*1000:,.0f}k", "wl_W12": mn(V["wl_W12"], 1),
    "wl_F2": num(V["wl_F2"]), "wl_F5": f"{V['wl_F5']:.2f}", "wl_F7": f"{V['wl_F7']:.2f}", "wl_F10": num(V["wl_F10"]), "wl_F11": num(V["wl_F11"]), "wl_Q2": f"{V['wl_Q2']:.0f}x", "wl_Q3": bn(V["wl_Q3"] / 1000, 2),
    "tco_T3": usd(V["tco_T3"]), "tco_T5": f"{V['tco_T5']:.1f}", "tco_T6": usd(V["tco_T6"]), "tco_T7": usd(V["tco_T7"]), "tco_T8": usd(V["tco_T8"]), "tco_T10": usd(V["tco_T10"]), "tco_T12": usd(V["tco_T12"]), "tco_T13": pct(V["tco_T13"]), "tco_T14": pct(V["tco_T14"]),
    "tco_T16": num(V["tco_T16"]), "tco_T17": usd(V["tco_T17"], 2), "tco_T18": usd(V["tco_T18"], 2), "tco_U1": usd(V["tco_U1"]), "tco_U2": usd(V["tco_U2"]), "tco_U5": pct(V["tco_U5"]), "tco_U6": pct(V["tco_U6"]), "tco_U7": pct(V["tco_U7"]), "tco_U8": usd(V["tco_U8"]),
    "tco_S1": mn(V["tco_S1"] / 1e6, 1), "tco_S2": num(V["tco_S2"]), "tco_S4": mn(V["tco_S4"], 1), "tco_S5": f"{V['tco_S5']:.2f}x",
    "val_VA3": f"{V['val_VA3']:.0f}x", "val_VA4": f"{V['val_VA4']*100:.1f} cents", "val_VA6": mn(V["val_VA6"]), "val_VA7": f"{V['val_VA7']*100:.2f}%", "val_VA8": mn(V["val_VA8"]), "val_VA9": mn(V["val_VA9"], 1), "val_VA10": f"{V['val_VA10']*12:.0f} months",
    "val_VB3": f"{V['val_VB3']:.1f}x", "val_VB4": bn(V["val_VB4"], 1), "val_VB5": pct(V["val_VB5"]), "val_VB6": mn(V["val_VB6"]), "val_VB7": f"{V['val_VB7']*100:.1f}%", "val_VB9": pct(V["val_VB9"]),
    "val_VC3": num(V["val_VC3"]), "val_VC5": num(V["val_VC5"]), "val_VC6": pct(V["val_VC6"]), "val_VC7": mn(V["val_VC7"], 1), "val_VC8": mn(V["val_VC8"]), "val_VC9": mn(V["val_VC9"], 1), "val_VC10": mn(V["val_VC10"]),
    "val_VD2": mn(V["val_VD2"] * 1000), "val_VD3": pct(V["val_VD3"], 1),
    "pr_P1": usd(V["pr_P1"], 2), "pr_P2": usd(V["pr_P2"], 2), "pr_P7": f"{V['pr_P7']:.1f}x", "pr_M1": f"{V['pr_M1']/1e6:.0f} million", "pr_M2": f"${V['pr_M2']:.2f}", "pr_M3": f"${V['pr_M3']:.2f}", "pr_M4": f"${V['pr_M4']:.2f}", "pr_M8": f"{V['pr_M8']:.0f}x",
    "pr_G3": usd(V["pr_G3"]), "pr_G4": pct(V["pr_G4"]), "pr_G5": pct(V["pr_G5"]), "pr_G6": usd(V["pr_G6"]), "pr_G7": usd(V["pr_G7"]), "pr_G8": pct(V["pr_G8"]), "pr_G9": pct(V["pr_G9"]), "pr_H4": usd(V["pr_H4"]),
}
rep.update(CF)
out = tmpl
for kk, vv in rep.items():
    out = out.replace("{{" + kk + "}}", str(vv))
import re
left = re.findall(r"\{\{[a-zA-Z0-9_]+\}\}", out)
if left:
    raise SystemExit("unfilled placeholders: " + ", ".join(sorted(set(left))))
open(HERE + "index.html", "w").write(out)
print("wrote index.html", len(out), "bytes")
