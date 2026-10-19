"""Build the GPU-in-financial-services model workbook.

Every number that is not a formula lives on the Assumptions sheet with a source.
Formulas reference those cells. The 'Formula' column on each sheet spells out
the calculation in words so a reader can follow it without opening the cell.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/claude/ai-radar/research/fs-gpu/fs-gpu-model.xlsx"

F = "Arial"
BLUE = Font(name=F, color="0000FF", size=10)
BLACK = Font(name=F, color="000000", size=10)
GREEN = Font(name=F, color="008000", size=10)
BOLD = Font(name=F, bold=True, size=10)
H1 = Font(name=F, bold=True, size=14)
H2 = Font(name=F, bold=True, size=11, color="FFFFFF")
NOTE = Font(name=F, italic=True, size=9, color="555555")
HEADFILL = PatternFill("solid", fgColor="132033")
KEYFILL = PatternFill("solid", fgColor="FFFF00")
SECFILL = PatternFill("solid", fgColor="E6F2F0")
thin = Side(style="thin", color="DDDDDD")
BORDER = Border(bottom=thin)

USD = '$#,##0;($#,##0);-'
USD1 = '$#,##0.0;($#,##0.0);-'
USD2 = '$#,##0.00;($#,##0.00);-'
NUM = '#,##0;(#,##0);-'
NUM1 = '#,##0.0;(#,##0.0);-'
PCT = '0.0%;(0.0%);-'
MULT = '0.0x'

wb = Workbook()

# ---------------------------------------------------------------- helpers
def header(ws, row, cols, titles):
    for c, t in zip(cols, titles):
        cell = ws.cell(row=row, column=c, value=t)
        cell.font = H2
        cell.fill = HEADFILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")

def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = H1
    if sub:
        ws["A2"] = sub
        ws["A2"].font = NOTE

def widths(ws, w):
    for i, x in enumerate(w, start=1):
        ws.column_dimensions[get_column_letter(i)].width = x

# ---------------------------------------------------------------- README
ws = wb.active
ws.title = "README"
title(ws, "GPU demand in financial services: sizing, cost, value and pricing model",
      "Companion to esthercs.dev/research/fs-gpu. Built 19 October 2026 by Esther Suebsubanant. All figures are dated and sourced on the Assumptions sheet.")
readme = [
    ("How to read this workbook", ""),
    ("Blue text", "Hardcoded input. Every one has a source or is labelled as my assumption on the Assumptions sheet."),
    ("Black text", "Formula on the same sheet."),
    ("Green text", "Formula that pulls a value from another sheet."),
    ("Yellow fill", "Scenario lever you are expected to change."),
    ("Formula column", "Each computed row states its formula in words, so you can audit without clicking into cells."),
    ("", ""),
    ("Sheets", ""),
    ("Assumptions", "Every input, with unit, value, date read and source URL."),
    ("TopDown", "Three independent top-down paths to the 2026 financial services accelerated-compute market, then a reconciliation."),
    ("BottomUp", "Segment by segment: number of institutions, adoption, GPUs per adopter, price per installed GPU, refresh cycle. Gives installed base, annual spend and a 2030 view."),
    ("Workloads", "Compute-demand sanity checks: how many GPUs an employee copilot, a fraud scorer and a quant research cluster actually need."),
    ("TCO", "Five-year total cost of ownership of one 8-GPU Blackwell node on-premises versus renting the same GPUs in the cloud, with the breakeven utilisation."),
    ("Configs", "Four reference configurations (regulated pod, bank SuperPOD class, sovereign NVL72 AI factory, quant research factory) evaluated on specification, five-year cost, unit economics, workload fit and value fit."),
    ("Value", "What the compute buys: fraud prevented, cost taken out, risk-engine consolidation, productivity. Value-to-cost ratios per use case."),
    ("Pricing", "Effective price per GPU-hour and per million tokens across buying routes, NVIDIA's gross margin, and where pricing power sits."),
    ("Sources", "Numbered source list matching the article."),
    ("", ""),
    ("Units", "Dollars are US dollars. Money in millions unless the header says otherwise. Percentages are stored as fractions. Years are calendar years unless marked FY (NVIDIA fiscal year ends late January)."),
]
r = 4
for k, v in readme:
    ws.cell(row=r, column=1, value=k).font = BOLD if v == "" else BLACK
    ws.cell(row=r, column=2, value=v).font = BLACK
    ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    r += 1
widths(ws, [22, 120])

# ---------------------------------------------------------------- Assumptions
A = wb.create_sheet("Assumptions")
title(A, "Assumptions and inputs", "Blue = hardcoded input. Change a value here and every sheet recalculates. 'My assumption' means no public figure exists and the number is my estimate, flagged so you can argue with it.")
header(A, 4, [1, 2, 3, 4, 5, 6], ["Key", "Input", "Value", "Unit", "Read / dated", "Source"])
widths(A, [8, 62, 16, 16, 14, 110])

ASSUMP = {}  # key -> cell address on Assumptions
row = 5

def section(name):
    global row
    c = A.cell(row=row, column=2, value=name)
    c.font = BOLD
    for col in range(1, 7):
        A.cell(row=row, column=col).fill = SECFILL
    row += 1

def inp(key, label, value, unit, dated, source, fmt=None, key_lever=False):
    global row
    A.cell(row=row, column=1, value=key).font = NOTE
    A.cell(row=row, column=2, value=label).font = BLACK
    c = A.cell(row=row, column=3, value=value)
    c.font = BLUE
    if fmt:
        c.number_format = fmt
    if key_lever:
        c.fill = KEYFILL
    A.cell(row=row, column=4, value=unit).font = BLACK
    A.cell(row=row, column=5, value=dated).font = BLACK
    s = A.cell(row=row, column=6, value=source)
    s.font = NOTE
    s.alignment = Alignment(wrap_text=False)
    ASSUMP[key] = f"Assumptions!$C${row}"
    row += 1

# --- Top-down market inputs
section("Top-down market inputs")
inp("it_bank", "Banking and investment services enterprise IT spending, 2026", 857.5, "$B", "4Q25 forecast", "Gartner, Forecast: Enterprise IT Spending for the Banking and Investment Services Market, Worldwide, 2023-2029, 4Q25. https://www.gartner.com/en/documents/7431362", USD1)
inp("it_bank_g", "Growth of that spending in 2026", 0.095, "fraction", "4Q25", "Same Gartner document (9.5%).", PCT)
inp("it_world", "Worldwide IT spending, 2026", 6369, "$B", "27 Jul 2026", "Gartner press release, 27 Jul 2026. https://www.gartner.com/en/newsroom/press-releases/2026-07-27-gartner-forecasts-worldwide-it-spending-to-grow-14-point-2-percent-in-2026-totaling-6-point-37-trillion", USD)
inp("dcs_world", "Worldwide data center systems spending, 2026", 822, "$B", "27 Jul 2026", "Same Gartner release, Table 1 (2025: $506B; 2026 growth 62.5%).", USD)
inp("dcs_world_25", "Worldwide data center systems spending, 2025", 506, "$B", "27 Jul 2026", "Same Gartner release, Table 1.", USD)
inp("ai_infra_26", "Worldwide AI infrastructure spending, 2026 (AI-optimised IaaS, AI servers, network fabric, AI semiconductors, devices)", 1484.397, "$B", "16 Sep 2026", "Gartner press release, 16 Sep 2026, Table 1. https://www.gartner.com/en/newsroom/press-releases/2026-09-16-gartner-forecasts-worldwide-ai-spending-to-grow-49-point-5-percent-in-2026", USD1)
inp("ai_infra_25", "Worldwide AI infrastructure spending, 2025", 981.920, "$B", "16 Sep 2026", "Same Gartner release.", USD1)
inp("ai_infra_27", "Worldwide AI infrastructure spending, 2027", 1977.685, "$B", "16 Sep 2026", "Same Gartner release.", USD1)
inp("ai_total_26", "Worldwide AI spending, all segments, 2026", 2670.460, "$B", "16 Sep 2026", "Same Gartner release ($2.67T, +49.5%).", USD1)
inp("ai_iaas_26", "Worldwide AI-optimised IaaS spending, 2026", 42.276, "$B", "10 Aug 2026", "Gartner press release, 10 Aug 2026. https://www.gartner.com/en/newsroom/press-releases/2026-08-10-gartner-forecasts-worldwide-artificial-intelligence-optimized-iaas-spending-to-grow-96-percent-in-2026", USD1)
inp("ai_iaas_inf", "Share of AI-optimised IaaS spend that is inference, 2026", 0.55, "fraction", "10 Aug 2026", "Same Gartner release ($23.3B inference, $19B training).", PCT)
inp("idc_bank_24", "Banking AI spending, 2024 (IDC)", 31.3, "$B", "21 Aug 2024", "IDC, Worldwide AI and Generative AI Spending Industry Outlook. https://www.idc.com/resource-center/blog/idcs-worldwide-ai-and-generative-ai-spending-industry-outlook/", USD1)
inp("idc_bank_cagr", "Banking AI spending five-year CAGR (IDC, Americas 30%, EMEA 32%)", 0.30, "fraction", "21 Aug 2024", "Same IDC post. I use the Americas figure as the global rate.", PCT)
inp("idc_bank_share", "Banking share of global AI spending implied by IDC (31.3 / 235)", 0.133, "fraction", "21 Aug 2024", "Same IDC post: global AI market ~$235B in 2024. 31.3/235 = 13.3%.", PCT)
inp("idc_ai_infra_26", "Worldwide AI infrastructure spending, 2026 (IDC, narrower definition than Gartner)", 487, "$B", "2026", "IDC figure as summarised by Digital Applied, 2026 (2025: $318B, 2024: $153B). https://www.digitalapplied.com/blog/ai-spending-forecasts-2026-gartner-idc-stanford-compiled", USD)
inp("ent_share_dcs", "Enterprise (non-hyperscaler, non-AI-lab) share of worldwide data center systems spend, 2026", 0.28, "fraction", "My assumption", "Four US hyperscalers guide ~$733B of 2026 capex (State of AI Report 2026), much of it servers; neoclouds and labs add more. Enterprises are the minority of the $822B. 28%. Lever.", PCT, True)
inp("fs_share_dcs", "Financial services share of enterprise data center systems spend", 0.17, "fraction", "My assumption", "Banking and investment services are $857.5B of $6,369B worldwide IT (13.5%). Capital markets and payments are compute-heavy relative to their IT budgets, so I lift the share to 17% of the enterprise slice. Lever.", PCT, True)
inp("hw_share_ai", "Share of financial services AI budget that goes to infrastructure (hardware plus cloud compute)", 0.30, "fraction", "22 Jan 2026", "NVIDIA State of AI in Financial Services 2026: 30% of respondents said new investment goes to AI infrastructure. I use it as a budget share, which is a stretch of the survey question. Lever. https://blogs.nvidia.com/blog/ai-in-financial-services-survey-2026/", PCT, True)
inp("accel_share_dcs", "Share of data center systems spend that is accelerated (GPU) compute in 2026", 0.62, "fraction", "My assumption", "Gartner has data center systems up 62.5% in 2026 after 51.6% in 2025, attributing the growth to AI servers. Treating the 2024 base (~$334B) as largely non-accelerated and all growth since as accelerated gives (822-334)/822 = 59%. Rounded up to 62% for FS, which over-indexes on HPC. Lever.", PCT, True)
inp("nvda_share", "NVIDIA share of data center AI accelerator revenue, 2026", 0.80, "fraction", "2026", "Secondary estimates cluster at 75% to 85% (IDC 81%, Silicon Analysts ~75%). https://commandlinux.com/statistics/ai-gpu-market-share-nvidia-vs-amd-vs-intel/", PCT, True)
inp("nvda_dc_q2", "NVIDIA Data Center revenue, Q2 FY2027 (quarter ended 26 Jul 2026)", 89.0, "$B", "Aug 2026", "NVIDIA press release. https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-second-quarter-fiscal-2027", USD1)
inp("nvda_dc_q1", "NVIDIA Data Center revenue, Q1 FY2027", 75.2, "$B", "May 2026", "NVIDIA press release. https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-first-quarter-fiscal-2027", USD1)
inp("nvda_gm", "NVIDIA GAAP gross margin, Q2 FY2027", 0.75, "fraction", "Aug 2026", "Same Q2 FY2027 release (75.0%).", PCT)
inp("nvda_q3_guide", "NVIDIA Q3 FY2027 revenue outlook", 108.0, "$B", "Aug 2026", "Same release ($108.0B plus or minus 2%).", USD1)

# --- Bottom-up: institutions
section("Bottom-up: institution counts")
inp("n_gsib", "Global systemically important banks", 29, "institutions", "27 Nov 2025", "FSB 2025 G-SIB list. https://www.fsb.org/2025/11/2025-list-of-global-systemically-important-banks-g-sibs/", NUM)
inp("n_top100", "Other banks in the world's 100 largest by assets (100 less 29 G-SIBs)", 71, "institutions", "29 Apr 2026", "S&P Global Market Intelligence, The world's largest banks by assets, 2026 (21 China, 12 US, 8 Japan). https://www.spglobal.com/market-intelligence/en/news-insights/research/2026/04/the-worlds-largest-banks-by-assets-2026", NUM)
inp("n_regional", "Regional and mid-size banks with a credible AI programme (ranks 101 to 500 globally)", 400, "institutions", "My assumption", "S&P Capital IQ Pro covers 15,000 banks; FDIC alone has 4,238 insured institutions (2Q 2026). I take ranks 101-500 as the addressable tier for owned GPUs. https://www.fdic.gov/news/press-releases/2026/fdic-insured-institutions-reported-return-assets-137-percent-and-net", NUM)
inp("n_payments", "Payment networks, processors and large fintechs (Visa, Mastercard, Amex, PayPal, Stripe, Adyen, Revolut, Nubank, etc.)", 25, "institutions", "My assumption", "Count of global networks, acquirers and fintechs with transaction foundation models in production or announced (GTC 2026 named Mastercard, Revolut, Adyen).", NUM)
inp("n_quant", "Quant trading and market-making firms with owned GPU research clusters", 40, "institutions", "My assumption", "XTX (>25,000 GPUs), HRT (Blackwell AI factory), Jump (Vera Rubin), Jane Street, Citadel Securities, Two Sigma, DE Shaw, Qube, Millennium, Point72, DRW, SIG, Optiver, IMC, High-Flyer and peers. https://www.xtxmarkets.com/", NUM)
inp("n_hedge", "Large hedge funds and asset managers (>$20B AUM) buying GPUs beyond the quant tier", 150, "institutions", "My assumption", "Order of magnitude from Preqin-style rankings; most rent rather than own.", NUM)
inp("n_insurers", "Large insurers and reinsurers with owned AI infrastructure", 60, "institutions", "My assumption", "Allianz, AXA, Ping An, Munich Re, Swiss Re, Zurich, Chubb, AIG, Prudential, MetLife and peers. Underwriting, claims and catastrophe modelling.", NUM)
inp("n_exchanges", "Exchanges, clearing houses and market data firms", 15, "institutions", "My assumption", "LSEG (GPU-accelerated analytics with NVIDIA, Oct 2026), ICE, CME, Nasdaq, DTCC, Bloomberg, S&P Global, MSCI. https://blogs.nvidia.com/feed/", NUM)

# --- Bottom-up: adoption and intensity
section("Bottom-up: adoption and GPU intensity, 2026")
inp("adopt_large", "Share of large institutions (>1,000 staff) actively using AI", 0.76, "fraction", "Mar 2026", "NVIDIA State of AI in FS 2026 as summarised by R. K. Thakur: 76% of firms over 1,000 employees. https://ninza7.medium.com/nvidias-ai-in-finance-report-just-confirmed-what-most-people-aren-t-ready-to-hear-2f26de01186d", PCT)
inp("adopt_all", "Share of all financial firms actively using AI", 0.65, "fraction", "22 Jan 2026", "NVIDIA State of AI in FS 2026 (65%, up from 45%).", PCT)
inp("own_share", "Of adopters, share that own or dedicate GPUs (rest rent via cloud only)", 0.55, "fraction", "My assumption", "NVIDIA 2026 survey: 47% hybrid. CSA/Anjuna 2026: hybrid and multi-cloud dominant. I set owned-or-dedicated at 55% for large institutions. Lever.", PCT, True)
inp("gpu_gsib", "Installed GPUs per G-SIB adopter, 2026", 6000, "GPUs", "My assumption", "BNY Mellon's 2024 SuperPOD was 'dozens' of DGX H100 (8 GPUs each), i.e. 300-500 GPUs for one bank's first system, with more coming; JPMorgan runs a $19.8B tech budget with 230,000 LLM Suite users. A G-SIB fleet across training, fraud, risk and inference of 3,000-10,000 GPUs is consistent with that; JPMorgan's ~$2B annual AI investment with 30% to infrastructure would alone buy ~12,000 GPUs a year. Fleet mean 6,000. Lever.", NUM, True)
inp("gpu_top100", "Installed GPUs per other top-100 bank adopter", 1000, "GPUs", "My assumption", "One to two SuperPOD-class systems or equivalent cloud-dedicated capacity.", NUM, True)
inp("gpu_regional", "Installed GPUs per regional bank adopter", 64, "GPUs", "My assumption", "Four to eight nodes: an inference pod for document processing and a fraud model.", NUM, True)
inp("gpu_payments", "Installed GPUs per payments adopter", 2000, "GPUs", "My assumption", "Transaction foundation models trained on hundreds of millions to 125B transactions (Mastercard) plus real-time scoring at up to 76,000 TPS (Visa).", NUM, True)
inp("gpu_quant", "Installed GPUs per quant firm", 4000, "GPUs", "My assumption", "XTX discloses >25,000; High-Flyer 10,000 A100; HRT and Jump build AI factories; Jane Street, Citadel Securities and Two Sigma are large; the tail (DRW, SIG, Optiver, IMC, Qube, Millennium pods) runs 500-3,000. Fleet mean across 40 firms 4,000. Lever.", NUM, True)
inp("gpu_hedge", "Installed or dedicated GPUs per large hedge fund / asset manager adopter", 150, "GPUs", "My assumption", "Mostly rented; dedicated reservations counted.", NUM, True)
inp("gpu_insurer", "Installed GPUs per insurer adopter", 200, "GPUs", "My assumption", "Claims, underwriting and catastrophe modelling; Allianz Partners fraud detection 1.5x (WEF 2026).", NUM, True)
inp("gpu_exchange", "Installed GPUs per exchange / market data adopter", 400, "GPUs", "My assumption", "LSEG, Bloomberg and S&P Global run market analytics and LLMs at scale.", NUM, True)
inp("price_gpu_sys", "System-level price per installed GPU (server, networking, storage share), 2026 blend of Hopper and Blackwell", 48000, "$ per GPU", "Sep-Oct 2026", "DGX B200 8-GPU $280k-$320k OEM anchor = $35k-$40k/GPU; HGX B200 $400k-$500k = $50k-$62k/GPU; GB200 NVL72 $2M-$3M = $28k-$42k/GPU; H100 server $250k-$320k = $31k-$40k/GPU. B200 chips alone are $45k-$50k on the street (Oct 2026). Blend $48k per installed GPU incl. server, networking and storage share, reflecting Hopper still in fleets. https://tech-insider.org/nvidia-blackwell-gpu-pricing/", USD, True)
inp("refresh_yrs", "Refresh cycle for owned GPUs", 5, "years", "12 Aug 2026", "NVIDIA says A100 (2020) remains in commercial use six years on and economic life is heading toward a decade; enterprises depreciate over 5-6 years. 5-year cycle. https://blogs.nvidia.com/blog/nvidia-ai-factory-compute/", NUM)
inp("fleet_growth", "Annual growth in installed GPU base, 2026 to 2030", 0.30, "fraction", "My assumption", "IDC banking AI spend CAGR 30%; Gartner AI infrastructure 2026-27 growth 33%. Lever.", PCT, True)
inp("rent_share_spend", "Cloud-rented GPU spend as a share of owned-GPU annual spend", 0.60, "fraction", "My assumption", "Gartner: AI-optimised IaaS is $42B against $822B data center systems worldwide (5%), but FS rents a larger share because regulation pushes workloads to dedicated cloud. Lever.", PCT, True)

# --- Workloads
section("Workload sanity checks")
inp("jpm_users", "JPMorgan employees on LLM Suite", 230000, "users", "1 Jul 2026", "Forbes / Bernard Marr. https://www.forbes.com/sites/bernardmarr/2026/07/01/how-jpmorgan-chase-is-building-the-ai-powered-bank-of-the-future/", NUM)
inp("tok_user_day", "Tokens per active user per working day (in plus out)", 60000, "tokens", "2026", "iternal.ai Token Usage Guide 2026: business analyst 300K-2M tokens/month (~15k-100k/day). 60k/day midpoint. https://iternal.ai/token-usage-guide", NUM, True)
inp("active_share", "Share of enrolled users active on a given day", 0.5, "fraction", "My assumption", "Enterprise copilot daily active rates.", PCT)
inp("peak_factor", "Peak-to-average ratio for interactive load", 3, "x", "My assumption", "Office hours concentrate load.", MULT)
inp("tps_gpu_70b", "Tokens per second per Blackwell GPU serving a 70B-class model at 55 tokens/s per user", 2500, "tokens/s", "16 Oct 2026", "Derived from NVIDIA's $0.02 per million tokens claim for GPT-OSS-120B on Blackwell with TensorRT-LLM; implies several thousand tokens/s per GPU at ~$6/GPU-hour. Conservative 2,500. https://www.nvidia.com/en-us/data-center/dgx-b200/", NUM, True)
inp("util_inf", "Target utilisation of inference GPUs", 0.6, "fraction", "My assumption", "Headroom for latency SLAs.", PCT)
inp("visa_txn_yr", "Visa transactions per year", 269e9, "transactions", "2026", "Visa, AI fraud detection page (over 269 billion). https://corporate.visa.com/en/solutions/visa-protect/insights/ai-fraud-detection.html", NUM)
inp("visa_tps_peak", "Visa peak capacity, transactions per second", 76000, "TPS", "4 Oct 2024", "Bloomberg Professional, Visa and Mastercard spur gen-AI adoption. https://www.bloomberg.com/professional/insights/artificial-intelligence/visa-mastercard-spur-gen-ai-adoption-among-payment-companies/", NUM)
inp("flop_txn", "FLOPs per transaction scored by a transaction foundation model (forward pass, ~1B parameter model, 2 FLOPs per parameter)", 2e9, "FLOPs", "My assumption", "Mastercard's model is trained on ~125B transactions a year and Revolut's on Hopper with Nemotron; a ~1B-parameter encoder at ~2 GFLOPs per scored transaction is a reasonable production size.", NUM)
inp("tfm_params", "Parameters of the transaction foundation model being trained", 1e10, "parameters", "My assumption", "Mastercard and Visa describe 'large transaction models' trained on 100B+ transactions; 10B parameters is a plausible production size. Scoring uses a distilled ~1B model (flop_txn).", NUM, True)
inp("pflops_gpu_fp8", "Effective FP8 throughput per Blackwell GPU in production", 2.0, "PFLOPS", "16 Oct 2026", "DGX B200: 72 PFLOPS FP8 across 8 GPUs = 9 PFLOPS peak per GPU; real inference achieves 20-30% of peak. 2 PFLOPS effective. https://www.nvidia.com/en-us/data-center/dgx-b200/", NUM1)
inp("xtx_gpus", "XTX Markets GPUs in research cluster", 25000, "GPUs", "Oct 2026", "XTX Markets homepage. https://www.xtxmarkets.com/", NUM)
inp("xtx_capex", "XTX Markets data center investment in Finland", 1.0, "$B", "2024-2025", "Bloomberg via Yahoo Finance: more than EUR 1 billion for five data centers. https://finance.yahoo.com/news/billionaire-alex-gerko-xtx-build-050000389.html", USD1)
inp("hrt_speedup", "HRT research iteration speedup on Blackwell vs Hopper", 1.6, "x", "17 Mar 2026", "NVIDIA GTC 2026 live updates. https://resources.nvidia.com/en-us-financial-services-industry/hudson-river", MULT)
inp("hrt_pue", "HRT Lefdal Mine Data Centers PUE", 1.15, "PUE", "17 Mar 2026", "Same source (1.15 or lower).", '0.00')

# --- TCO
section("TCO: one 8-GPU Blackwell node, on-premises vs cloud")
inp("node_price", "8-GPU Blackwell (HGX/DGX B200 class) system price", 420000, "$", "Apr-Oct 2026", "DGX B200 OEM anchor $280k-$320k (Apr 2026) predates the autumn price rise; HGX B200 server $400k-$500k (GPU Smith, Sep 2026); Broadberry DGX B200 config $562k (21 Sep 2026); 8 x $47.5k street B200 = $380k in chips alone. $420k. Lever. https://tech-insider.org/nvidia-blackwell-gpu-pricing/", USD, True)
inp("node_netstor", "Networking, storage and installation as a share of system price", 0.25, "fraction", "My assumption", "InfiniBand or Spectrum-X fabric, NVMe storage, racks, PDUs, integration. SuperPOD reference architectures run 20-30%.", PCT)
inp("node_kw", "DGX B200 maximum power draw", 14.3, "kW", "16 Oct 2026", "NVIDIA DGX B200 specifications (~14.3 kW max). https://www.nvidia.com/en-us/data-center/dgx-b200/", NUM1)
inp("node_load", "Average power as a share of maximum", 0.75, "fraction", "My assumption", "Training runs near max; inference fleets idle more.", PCT)
inp("pue", "Power usage effectiveness of the facility", 1.3, "PUE", "My assumption", "Enterprise colocation 1.2-1.5; HRT's Lefdal facility 1.15. Lever.", '0.00', True)
inp("elec_price", "Industrial electricity price, US average", 0.0977, "$ per kWh", "Jul 2026", "EIA Electricity Monthly Update, end use, July 2026: industrial 9.77 c/kWh. https://www.eia.gov/electricity/monthly/update/end-use.php", '$0.0000')
inp("colo_kw_mo", "Colocation rate, primary North American markets, 250-500 kW", 204, "$ per kW per month", "mid-2026", "Metro Colo Advisory calculation from CBRE (~$204); CBRE H2 2025 $196.25; Chicago $200-230; N. Virginia $190-235. https://www.cbre.com/insights/reports/global-data-center-trends-2026", USD)
inp("sw_gpu_yr", "NVIDIA AI Enterprise and support, per GPU per year", 4500, "$ per GPU per year", "My assumption", "List pricing has run $4,500 per GPU per year for AI Enterprise; DGX systems bundle it for the first years. Lever.", USD, True)
inp("staff_node_yr", "Operations staff cost allocated per node per year", 15000, "$ per node per year", "My assumption", "One platform engineer ($250k loaded) per ~16 nodes.", USD)
inp("yrs", "Ownership period", 5, "years", "", "Matches refresh cycle.", NUM)
inp("cloud_b200_hr", "Cloud B200 on-demand price per GPU-hour (hyperscaler-grade)", 6.99, "$ per GPU-hour", "Sep-Oct 2026", "Northflank $6.99 single GPU; range $5.99 (Hyperbolic) to $16.11 (Google Cloud); NVIDIA cites $5.30-$7.05. https://northflank.com/blog/how-much-does-an-nvidia-b200-gpu-cost", USD2, True)
inp("cloud_b200_res", "Cloud B200 reserved (1-3 year) price per GPU-hour", 4.75, "$ per GPU-hour", "2026", "Reserved discounts of 30-35% off on-demand are typical; Spheron spot $5.22. https://www.spheron.network/blog/nvidia-h200-vs-b200-vs-gb200/", USD2, True)
inp("cloud_h100_hr", "Cloud H100 on-demand median price per GPU-hour", 2.70, "$ per GPU-hour", "Jun 2026", "NVIDIA, AI Factory Compute Is Becoming an Investable Asset Class (median $2.70 in June 2026, up from $2.00 in Oct 2025). https://blogs.nvidia.com/blog/nvidia-ai-factory-compute/", USD2)
inp("cloud_h100_1y", "Cloud H100 one-year rental price per GPU-hour", 2.35, "$ per GPU-hour", "Mar 2026", "Same NVIDIA post ($1.70 Oct 2025 to $2.35 Mar 2026).", USD2)
inp("util_own", "Utilisation of owned GPUs (share of hours busy)", 0.65, "fraction", "My assumption", "Run:ai claims 5x utilisation gains from scheduling; banks without it sit at 30-40%. Lever.", PCT, True)
inp("hours_yr", "Hours per year", 8760, "hours", "", "365 x 24.", NUM)
inp("resid_val", "Residual value of hardware at end of year 5 as share of price", 0.10, "fraction", "12 Aug 2026", "NVIDIA offers residual-value support up to 25% of an opportunity in its financing platforms; secondary H100 prices have held. 10% is conservative.", PCT)

# --- Configurations
section("System configuration inputs")
inp("b200_mem", "HBM3e memory per B200 GPU", 180, "GB", "16 Oct 2026", "DGX B200: 1,440 GB across 8 GPUs. https://www.nvidia.com/en-us/data-center/dgx-b200/", NUM)
inp("b300_mem", "HBM3e memory per B300 GPU", 288, "GB", "2026", "Blackwell Ultra (B300) public specification. My figure.", NUM)
inp("node_fp8", "FP8 Tensor PFLOPS per 8-GPU B200 node (with sparsity, as NVIDIA quotes)", 72, "PFLOPS", "16 Oct 2026", "DGX B200 page: FP8 72 PFLOPS. https://www.nvidia.com/en-us/data-center/dgx-b200/", NUM)
inp("nvl72_fp8_mult", "GB200 NVL72 FP8 throughput relative to one 8-GPU node", 9, "x", "2026", "72 GPUs / 8 GPUs. NVLink domain across all 72 GPUs is the architectural difference, not the FLOPS.", MULT)
inp("gb300_price", "GB300 NVL72 rack price, midpoint of estimates", 3.5, "$M", "Oct 2026", "$3M-$4M estimates. https://tech-insider.org/nvidia-blackwell-gpu-pricing/", USD1)
inp("gb300_kw", "GB300 NVL72 rack power", 135, "kW", "2026", "Reported 120-140 kW liquid-cooled. My figure.", NUM)
inp("hgx_air_price", "Air-cooled HGX B200 server, 8 GPUs, from an ODM (research-factory buyer, no DGX software)", 380000, "$", "Sep 2026", "GPU Smith prices a configured HGX B200 at $400k-$500k; large quant buyers negotiate toward the chip price (8 x $47.5k = $380k). Lever.", USD, True)
inp("rack_fabric_nvl", "Networking, storage and installation share for NVL72 racks", 0.15, "fraction", "My assumption", "Scale-up NVLink is inside the rack price; only scale-out fabric and storage are extra.", PCT)
inp("liquid_premium", "Facility premium for liquid-cooled racks, share of colocation rate", 0.20, "fraction", "My assumption", "CBRE notes limited liquid-cooled inventory; Digital Realty Frankfurt offers liquid-cooled racks at a premium. 20%.", PCT)

# --- Value
section("Value inputs")
inp("visa_fraud", "Fraud prevented annually by Visa Advanced Authorization", 28, "$B", "4 Oct 2024", "Bloomberg Professional. https://www.bloomberg.com/professional/insights/artificial-intelligence/visa-mastercard-spur-gen-ai-adoption-among-payment-companies/", USD1)
inp("visa_ai_spend", "Visa spending on AI and data infrastructure over ten years", 3, "$B", "4 Oct 2024", "Same source (more than $3B).", USD1)
inp("mc_fraud_26", "Fraud prevented by Mastercard Decision Intelligence Pro in 2026 to date", 13, "$B", "2026", "Mastercard presentation summarised on YouTube; treat as indicative. https://www.youtube.com/watch?v=PWaGdoCnn5A", USD1)
inp("mc_txn_yr", "Mastercard payment transactions per year", 125e9, "transactions", "4 Oct 2024", "Bloomberg Professional (roughly 125 billion).", NUM)
inp("fraud_global", "Global financial impact of fraud, 2024", 485, "$B", "6 Feb 2026", "Mastercard Insights citing industry research. https://www.mastercard.com/us/en/news-and-trends/Insights/2026/ai-is-helping-banks-save-millions-by-transforming-payment-fraud-prevention.html", USD)
inp("fraud_org_loss", "Average annual payment fraud loss per organisation", 60, "$M", "2025", "Same Mastercard source (2025 research).", USD)
inp("fraud_saved_5y", "Fraud savings for organisations using AI more than five years", 4.3, "$M per year", "6 Feb 2026", "Same source ($4.3M vs $2.2M average).", USD1)
inp("revolut_prec", "Revolut fraud detection precision gain from transaction foundation model", 0.20, "fraction", "17 Mar 2026", "NVIDIA GTC 2026 (20%). https://resources.nvidia.com/en-us-financial-services-industry/hudson-river", PCT)
inp("jpm_savings", "JPMorgan annual savings attributed to AI", 2.0, "$B", "1 Jul 2026", "Forbes / Bernard Marr. https://www.forbes.com/sites/bernardmarr/2026/07/01/how-jpmorgan-chase-is-building-the-ai-powered-bank-of-the-future/", USD1)
inp("jpm_ai_spend", "JPMorgan annual AI investment", 2.0, "$B", "2026", "Olakai summary citing JPMorgan ($2B within $19.8B tech budget). https://olakai.ai/blog/jpmorgan-ai-measurement/", USD1)
inp("jpm_tech", "JPMorgan technology budget, 2026", 19.8, "$B", "2026", "JPMorganChase technology page. https://www.jpmorganchase.com/about/technology", USD1)
inp("jpm_eff", "Self-reported efficiency gain of LLM Suite users", 0.35, "fraction", "1 Jul 2026", "Forbes: 30-40%. Midpoint.", PCT)
inp("loaded_cost", "Loaded annual cost per knowledge worker", 150000, "$", "My assumption", "Blend of operations, technology and front-office staff at a large US bank.", USD)
inp("time_ai_share", "Share of a user's working time spent on tasks where the assistant applies", 0.25, "fraction", "My assumption", "Accenture: 73% of banking employees' time could be affected by genAI; only part is in use today. Lever.", PCT, True)
inp("murex_perf", "Murex XVA performance gain on NVIDIA GPUs vs CPU", 7, "x", "3 Feb 2026", "NVIDIA finance partners page. https://www.nvidia.com/en-us/industries/finance/partners/", MULT)
inp("murex_energy", "Murex XVA energy reduction vs CPU", 4, "x", "3 Feb 2026", "Same source (4x reduction).", MULT)
inp("cpu_server_price", "Dual-socket CPU risk-grid server price", 25000, "$", "My assumption", "High-core-count x86 node with memory.", USD)
inp("cpu_server_kw", "CPU server power", 1.0, "kW", "My assumption", "Dual socket under load.", NUM1)
inp("risk_grid_servers", "CPU servers in a large bank's XVA risk grid", 5000, "servers", "My assumption", "Large dealer grids run thousands to tens of thousands of cores.", NUM)
inp("swedbank", "Swedbank operating cost reduction from GPU fraud models (Hopsworks)", 10, "$M", "3 Feb 2026", "NVIDIA finance partners page (greater than $10M). https://www.nvidia.com/en-us/industries/finance/partners/", USD)
inp("adyen_speed", "Adyen foundation model inference speedup on NVIDIA", 195, "x", "17 Mar 2026", "NVIDIA GTC 2026.", MULT)

# --- Pricing
section("Pricing inputs")
inp("tok_cost_bw", "NVIDIA stated inference cost on Blackwell, GPT-OSS-120B, TensorRT-LLM", 0.02, "$ per million tokens", "16 Oct 2026", "DGX B200 page. https://www.nvidia.com/en-us/data-center/dgx-b200/", '$0.000')
inp("tok_cost_hopper", "Same on Hopper with vLLM", 0.09, "$ per million tokens", "16 Oct 2026", "Same page.", '$0.000')
inp("api_price_frontier", "Frontier API list price, blended input/output, per million tokens", 6.0, "$ per million tokens", "Oct 2026", "Typical blended list for a frontier model ($3 in / $15 out at 3:1). My blend.", USD2)
inp("api_price_open", "Hosted open-model API price per million tokens", 0.60, "$ per million tokens", "Oct 2026", "Typical hosted 70B-class open model price. My blend.", USD2)
inp("b200_chip", "B200 GPU standalone street price", 47500, "$", "2 Oct 2026", "IntuitionLabs via tech-insider: $45k-$50k, up from $30k-$40k. https://tech-insider.org/nvidia-blackwell-gpu-pricing/", USD)
inp("b300_chip", "B300 GPU standalone street price", 53000, "$", "13 Oct 2026", "Same source.", USD)
inp("nvl72_price", "GB200 NVL72 rack price, midpoint of estimates", 2.5, "$M", "Jul 2026", "$2M-$3M estimates. https://tech-insider.org/nvidia-blackwell-gpu-pricing/", USD1)
inp("nvl72_kw", "GB200 NVL72 rack power", 120, "kW", "2026", "Widely reported ~120 kW liquid-cooled rack. My figure.", NUM)
inp("fin_500", "Third-party capital NVIDIA's financing platforms aim to mobilise", 500, "$B", "12 Aug 2026", "NVIDIA / Jensen Huang blog with Apollo, BlackRock, Blackstone, Brookfield, Goldman Sachs, KKR. https://blogs.nvidia.com/blog/nvidia-ai-factory-compute/", USD)
inp("fin_resid", "NVIDIA residual-value support, share of an opportunity", 0.25, "fraction", "12 Aug 2026", "Same source (up to 25%).", PCT)

A.freeze_panes = "A5"

# ---------------------------------------------------------------- TopDown
T = wb.create_sheet("TopDown")
title(T, "Top-down: three paths to the 2026 financial services accelerated-compute market",
      "Each path starts from a published market figure and applies stated shares. Column D spells out the formula. $B unless noted.")
header(T, 4, [1, 2, 3, 4], ["Line", "Item", "Value", "Formula (in words) and source"])
widths(T, [6, 70, 18, 110])

def link(ws, key):
    return "=" + ASSUMP[key]

tr = 5
def trow(n, label, formula, words, fmt=USD1, font=BLACK, bold=False, fill=None):
    global tr
    T.cell(row=tr, column=1, value=n).font = NOTE
    c1 = T.cell(row=tr, column=2, value=label)
    c1.font = BOLD if bold else BLACK
    c = T.cell(row=tr, column=3, value=formula)
    c.font = font
    c.number_format = fmt
    if bold:
        c.font = Font(name=F, bold=True, size=10, color=font.color)
    w = T.cell(row=tr, column=4, value=words)
    w.font = NOTE
    w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 5):
            T.cell(row=tr, column=col).fill = fill
    tr += 1
    return f"C{tr-1}"

def tsec(name):
    global tr
    c = T.cell(row=tr, column=2, value=name)
    c.font = BOLD
    for col in range(1, 5):
        T.cell(row=tr, column=col).fill = SECFILL
    tr += 1

tsec("Path A: from the data center systems market (Gartner)")
a1 = trow("A1", "Banking and investment services IT spending, 2026", link(T, "it_bank"), "Gartner 4Q25 forecast: $857.5B, up 9.5%.", USD1, GREEN)
a2 = trow("A2", "Worldwide IT spending, 2026", link(T, "it_world"), "Gartner Jul 2026: $6,369B.", USD1, GREEN)
a3 = trow("A3", "Banking share of worldwide IT", f"={a1}/{a2}", "A1 / A2. Banks are about one dollar in seven of world IT.", PCT)
a4 = trow("A4", "Worldwide data center systems spending, 2026", link(T, "dcs_world"), "Gartner Jul 2026: $822B, up 62.5%.", USD1, GREEN)
a4b = trow("A5", "Enterprise share of data center systems (excludes hyperscalers, neoclouds, AI labs)", link(T, "ent_share_dcs"), "Assumption 28%: hyperscaler capex alone is ~$700B in 2026.", PCT, GREEN)
a4c = trow("A6", "Enterprise data center systems spend", f"={a4}*{a4b}", "A4 x A5.", USD1)
a5 = trow("A7", "Financial services share of enterprise data center systems", link(T, "fs_share_dcs"), "Assumption 17%, above the 13.5% IT share because capital markets and payments are compute-heavy.", PCT, GREEN)
a6 = trow("A8", "Financial services data center systems spend", f"={a4c}*{a5}", "A6 x A7.", USD1)
a7 = trow("A9", "Share of data center systems that is accelerated compute", link(T, "accel_share_dcs"), "Assumption 62%. See Assumptions for derivation from Gartner growth.", PCT, GREEN)
a8 = trow("A10", "Financial services accelerated-compute systems, 2026 (Path A)", f"={a6}*{a7}", "A8 x A9. Owned servers and racks, not cloud rental and not chips alone.", USD1, BLACK, True, KEYFILL)
a9 = trow("A11", "NVIDIA share of accelerators", link(T, "nvda_share"), "Assumption 80% (IDC 81%, Silicon Analysts ~75%).", PCT, GREEN)
a10 = trow("A12", "NVIDIA-addressable FS accelerated systems, 2026 (Path A)", f"={a8}*{a9}", "A10 x A11.", USD1, BLACK, True)

tsec("Path B: from the AI cloud market (Gartner AI-optimised IaaS)")
b1 = trow("B1", "Worldwide AI-optimised IaaS spending, 2026", link(T, "ai_iaas_26"), "Gartner Aug 2026: $42.3B, up 96%. The GPU-rental market.", USD1, GREEN)
b2 = trow("B2", "Financial services share of AI cloud spend", link(T, "idc_bank_share"), "IDC: banking was 13.3% of global AI spend in 2024 and the #2 industry. Applied to AI IaaS.", PCT, GREEN)
b3 = trow("B3", "Financial services AI cloud rental, 2026", f"={b1}*{b2}", "B1 x B2.", USD1)
b4 = trow("B4", "Rented share of owned spend (lever)", link(T, "rent_share_spend"), "Assumption 60%: for every $1 FS spends on owned GPU systems it rents $0.60 in the cloud.", PCT, GREEN)
b5 = trow("B5", "Implied owned GPU systems", f"={b3}/{b4}", "B3 / B4.", USD1)
b6 = trow("B6", "Financial services GPU compute, owned plus rented, 2026 (Path B)", f"={b3}+{b5}", "B3 + B5.", USD1, BLACK, True, KEYFILL)
b7 = trow("B7", "Context: worldwide AI infrastructure incl. devices and hyperscalers (Gartner)", link(T, "ai_infra_26"), "Gartner Sep 2026: $1,484B. Not used in the path because it includes AI PCs, phones and hyperscaler buildout.", USD1, GREEN)

tsec("Path C: from the banking AI budget (IDC, grown forward)")
c1 = trow("C1", "Banking AI spending, 2024", link(T, "idc_bank_24"), "IDC Aug 2024: $31.3B.", USD1, GREEN)
c2 = trow("C2", "CAGR applied", link(T, "idc_bank_cagr"), "IDC Americas five-year CAGR 30%.", PCT, GREEN)
c3 = trow("C3", "Banking AI spending, 2026", f"={c1}*(1+{c2})^2", "C1 x (1 + C2)^2. Two years of growth.", USD1)
c4 = trow("C4", "Uplift for non-bank financial services (insurance, capital markets, payments) over banking", 0.45, "Assumption: IDC's banking line excludes insurers, exchanges, asset managers and the quant firms, which carry the largest GPU fleets. +45%.", PCT, BLUE)
c5 = trow("C5", "Financial services AI spending, 2026", f"={c3}*(1+{c4})", "C3 x (1 + C4).", USD1)
c6 = trow("C6", "Share of AI budget spent on infrastructure", link(T, "hw_share_ai"), "NVIDIA 2026 survey: 30% of new investment to AI infrastructure.", PCT, GREEN)
c7 = trow("C7", "Financial services AI infrastructure, 2026 (Path C)", f"={c5}*{c6}", "C5 x C6.", USD1, BLACK, True, KEYFILL)

tsec("Reconciliation")
r1 = trow("R1", "Path A: accelerated systems from the IT budget", f"={a8}", "", USD1)
r2 = trow("R2", "Path B: GPU compute from the AI cloud market", f"={b6}", "", USD1)
r3 = trow("R3", "Path C: AI infrastructure from the AI budget", f"={c7}", "", USD1)
r4 = trow("R4", "Average of the three paths", f"=AVERAGE({r1},{r2},{r3})", "AVERAGE(R1:R3). Path A is owned systems only; B and C include cloud. The spread is informative, not an error.", USD1, BLACK, True, KEYFILL)
r5 = trow("R5", "Spread (max minus min) as share of average", f"=(MAX({r1},{r2},{r3})-MIN({r1},{r2},{r3}))/{r4}", "(MAX - MIN) / R4.", PCT)
r6 = trow("R6", "Bottom-up total for comparison (owned plus rented, 2026)", "=BottomUp!C32", "Pulled from BottomUp sheet, line 'Total FS GPU spend 2026 incl. rented'.", USD1, GREEN, True)
r7 = trow("R7", "Bottom-up as share of top-down average", f"={r6}/{r4}", "R6 / R4. Below 100% means the segment build does not reach what the budget view implies.", PCT)
r7b = trow("R7b", "Gap: top-down average minus bottom-up", f"={r4}-{r6}", "R4 - R6.", USD1)
r7c = trow("R7c", "Share of top-down 'infrastructure' that is not GPU systems (storage, networking, CPU servers, data platforms, cloud services around the GPUs)", 0.35, "Assumption: in an AI system bill of materials the GPU server is 60-70%; the rest is fabric, storage, CPU nodes and software. 35% of top-down infrastructure is therefore outside the bottom-up GPU count.", PCT, BLUE)
r7d = trow("R7d", "Top-down average restated to GPU systems only", f"={r4}*(1-{r7c})", "R4 x (1 - R7c).", USD1)
r7e = trow("R7e", "Bottom-up as share of restated top-down", f"={r6}/{r7d}", "R6 / R7d. What remains is the amount of undisclosed GPU intensity at the big banks and quant firms, or over-statement in the budget figures. Both are plausible.", PCT, BLACK, True, KEYFILL)
r8 = trow("R8", "FS as share of NVIDIA Data Center revenue run-rate (Path A NVIDIA-addressable / annualised Q2 FY27)", f"={a10}/({ASSUMP['nvda_dc_q2']}*4)", "A12 / (Q2 FY27 Data Center revenue x 4). Order of magnitude of FS inside NVIDIA's ~$356B run-rate. NVIDIA does not disclose a vertical split.", PCT)

T.freeze_panes = "A5"

# ---------------------------------------------------------------- BottomUp
B = wb.create_sheet("BottomUp")
title(B, "Bottom-up: GPUs installed and bought by financial services segment, 2026 and 2030",
      "Installed base = institutions x adoption x owned share x GPUs per adopter. Annual spend = installed base / refresh years x price per GPU, plus growth capex. $M unless noted.")
header(B, 4, list(range(1, 13)), ["Segment", "Institutions", "AI adoption", "Own or dedicate GPUs", "Adopters owning GPUs", "GPUs per adopter", "Installed GPUs 2026", "Price per installed GPU ($)", "Installed value ($M)", "Replacement spend 2026 ($M)", "Growth spend 2026 ($M)", "Total owned spend 2026 ($M)"])
widths(B, [44, 13, 12, 14, 14, 13, 16, 16, 16, 18, 16, 18])
B.row_dimensions[4].height = 45

segments = [
    ("G-SIBs", "n_gsib", "adopt_large", "gpu_gsib"),
    ("Other top-100 banks", "n_top100", "adopt_large", "gpu_top100"),
    ("Regional and mid-size banks (ranks 101-500)", "n_regional", "adopt_all", "gpu_regional"),
    ("Payment networks, processors, large fintechs", "n_payments", "adopt_large", "gpu_payments"),
    ("Quant trading and market-making firms", "n_quant", "adopt_large", "gpu_quant"),
    ("Large hedge funds and asset managers", "n_hedge", "adopt_large", "gpu_hedge"),
    ("Large insurers and reinsurers", "n_insurers", "adopt_large", "gpu_insurer"),
    ("Exchanges, clearing, market data", "n_exchanges", "adopt_large", "gpu_exchange"),
]
br = 5
first = br
for name, nk, ak, gk in segments:
    B.cell(row=br, column=1, value=name).font = BLACK
    B.cell(row=br, column=2, value="=" + ASSUMP[nk]).font = GREEN
    B.cell(row=br, column=2).number_format = NUM
    if nk == "n_quant":
        B.cell(row=br, column=3, value=1.0).font = BLUE
    else:
        B.cell(row=br, column=3, value="=" + ASSUMP[ak]).font = GREEN
    B.cell(row=br, column=3).number_format = PCT
    # quant firms own by definition
    if nk == "n_quant":
        B.cell(row=br, column=4, value=1.0).font = BLUE
    else:
        B.cell(row=br, column=4, value="=" + ASSUMP["own_share"]).font = GREEN
    B.cell(row=br, column=4).number_format = PCT
    B.cell(row=br, column=5, value=f"=B{br}*C{br}*D{br}").number_format = NUM1
    B.cell(row=br, column=6, value="=" + ASSUMP[gk]).font = GREEN
    B.cell(row=br, column=6).number_format = NUM
    B.cell(row=br, column=7, value=f"=E{br}*F{br}").number_format = NUM
    B.cell(row=br, column=8, value="=" + ASSUMP["price_gpu_sys"]).font = GREEN
    B.cell(row=br, column=8).number_format = USD
    B.cell(row=br, column=9, value=f"=G{br}*H{br}/1000000").number_format = USD1
    B.cell(row=br, column=10, value=f"=I{br}/{ASSUMP['refresh_yrs']}").number_format = USD1
    B.cell(row=br, column=11, value=f"=I{br}*{ASSUMP['fleet_growth']}").number_format = USD1
    B.cell(row=br, column=12, value=f"=J{br}+K{br}").number_format = USD1
    for col in (5, 7, 9, 10, 11, 12):
        B.cell(row=br, column=col).font = BLACK
    br += 1
last = br - 1
B.cell(row=br, column=1, value="Total").font = BOLD
for col in (2, 5, 7, 9, 10, 11, 12):
    L = get_column_letter(col)
    c = B.cell(row=br, column=col, value=f"=SUM({L}{first}:{L}{last})")
    c.font = BOLD
    c.number_format = B.cell(row=last, column=col).number_format
    c.fill = KEYFILL
total_row = br
br += 2

# Formula legend
legend = [
    ("Adopters owning GPUs", "Institutions x AI adoption x Own-or-dedicate share"),
    ("Installed GPUs 2026", "Adopters owning GPUs x GPUs per adopter"),
    ("Installed value ($M)", "Installed GPUs x Price per installed GPU / 1,000,000"),
    ("Replacement spend 2026", "Installed value / Refresh cycle (years). Steady-state refresh of the fleet."),
    ("Growth spend 2026", "Installed value x Fleet growth rate. New capacity added this year."),
    ("Total owned spend 2026", "Replacement + Growth"),
]
B.cell(row=br, column=1, value="Formulas").font = BOLD
br += 1
for k, v in legend:
    B.cell(row=br, column=1, value=k).font = BLACK
    B.cell(row=br, column=2, value=v).font = NOTE
    br += 1
br += 1

# Summary block starting at a fixed row for cross-sheet references: ensure row 32 is the incl-rented total.
# Current br should be <= 25. Pad to row 27.
while br < 27:
    br += 1
B.cell(row=br, column=1, value="Summary 2026").font = BOLD
br += 1  # 28
B.cell(row=br, column=1, value="Installed GPUs, financial services, 2026").font = BLACK
B.cell(row=br, column=3, value=f"=G{total_row}").number_format = NUM
B.cell(row=br, column=4, value="Sum of column G.").font = NOTE
inst_cell = f"C{br}"
br += 1  # 29
B.cell(row=br, column=1, value="Installed value ($M)").font = BLACK
B.cell(row=br, column=3, value=f"=I{total_row}").number_format = USD1
br += 1  # 30
B.cell(row=br, column=1, value="Owned GPU spend 2026 ($M)").font = BLACK
B.cell(row=br, column=3, value=f"=L{total_row}").number_format = USD1
owned_cell = f"C{br}"
br += 1  # 31
B.cell(row=br, column=1, value="Cloud-rented GPU spend 2026 ($M)").font = BLACK
B.cell(row=br, column=3, value=f"={owned_cell}*{ASSUMP['rent_share_spend']}").number_format = USD1
B.cell(row=br, column=4, value="Owned spend x rented-share lever (60%).").font = NOTE
rent_cell = f"C{br}"
br += 1  # 32
assert br == 32, br
B.cell(row=br, column=1, value="Total FS GPU spend 2026 incl. rented ($B)").font = BOLD
c = B.cell(row=br, column=3, value=f"=({owned_cell}+{rent_cell})/1000")
c.number_format = USD1
c.font = BOLD
c.fill = KEYFILL
B.cell(row=br, column=4, value="(Owned + Rented) / 1000. Feeds TopDown reconciliation R6.").font = NOTE
br += 1  # 33
B.cell(row=br, column=1, value="NVIDIA share of that spend ($B)").font = BLACK
B.cell(row=br, column=3, value=f"=C32*{ASSUMP['nvda_share']}").number_format = USD1
br += 1
B.cell(row=br, column=1, value="Quant and payments share of installed GPUs").font = BLACK
B.cell(row=br, column=3, value=f"=(G{first+3}+G{first+4})/G{total_row}").number_format = PCT
B.cell(row=br, column=4, value="(payments GPUs + quant GPUs) / total. Shows how concentrated the fleet is outside the banks.").font = NOTE
br += 2

# 2030 projection
B.cell(row=br, column=1, value="Projection 2026 to 2030").font = BOLD
br += 1
header(B, br, [1, 2, 3, 4, 5, 6], ["Year", "Installed GPUs", "Installed value ($M)", "Replacement ($M)", "Growth ($M)", "Owned spend ($M)"])
br += 1
yr0 = br
for i, y in enumerate([2026, 2027, 2028, 2029, 2030]):
    B.cell(row=br, column=1, value=str(y)).font = BLACK
    if i == 0:
        B.cell(row=br, column=2, value=f"={inst_cell}").number_format = NUM
    else:
        B.cell(row=br, column=2, value=f"=B{br-1}*(1+{ASSUMP['fleet_growth']})").number_format = NUM
    B.cell(row=br, column=3, value=f"=B{br}*{ASSUMP['price_gpu_sys']}/1000000").number_format = USD1
    B.cell(row=br, column=4, value=f"=C{br}/{ASSUMP['refresh_yrs']}").number_format = USD1
    B.cell(row=br, column=5, value=f"=C{br}*{ASSUMP['fleet_growth']}").number_format = USD1
    B.cell(row=br, column=6, value=f"=D{br}+E{br}").number_format = USD1
    br += 1
B.cell(row=br, column=1, value="Formulas: GPUs(t) = GPUs(t-1) x (1 + fleet growth). Value = GPUs x price per GPU. Replacement = value / refresh years. Growth = value x growth rate. Price per GPU held flat: Blackwell to Rubin raises price per GPU but tokens per dollar fall faster, so dollars per installed GPU is roughly stable.").font = NOTE
B.freeze_panes = "B5"

# ---------------------------------------------------------------- Workloads
W = wb.create_sheet("Workloads")
title(W, "Workload sanity checks: how many GPUs does each use case actually need?",
      "Three builds from first principles. The point is the ratio between them: inference for staff and payments is small; research and training are where the fleets live.")
header(W, 4, [1, 2, 3, 4], ["Line", "Item", "Value", "Formula (in words)"])
widths(W, [6, 70, 18, 110])
wr = 5
def wrow(n, label, formula, words, fmt=NUM, font=BLACK, bold=False, fill=None):
    global wr
    W.cell(row=wr, column=1, value=n).font = NOTE
    W.cell(row=wr, column=2, value=label).font = BOLD if bold else BLACK
    c = W.cell(row=wr, column=3, value=formula)
    c.font = Font(name=F, bold=bold, size=10, color=font.color)
    c.number_format = fmt
    w = W.cell(row=wr, column=4, value=words)
    w.font = NOTE
    w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 5):
            W.cell(row=wr, column=col).fill = fill
    wr += 1
    return f"C{wr-1}"
def wsec(name):
    global wr
    W.cell(row=wr, column=2, value=name).font = BOLD
    for col in range(1, 5):
        W.cell(row=wr, column=col).fill = SECFILL
    wr += 1

wsec("1. Employee assistant at a G-SIB (JPMorgan LLM Suite scale)")
w1 = wrow("W1", "Enrolled users", "=" + ASSUMP["jpm_users"], "230,000 (Forbes, Jul 2026).", NUM, GREEN)
w2 = wrow("W2", "Daily active share", "=" + ASSUMP["active_share"], "Assumption 50%.", PCT, GREEN)
w3 = wrow("W3", "Tokens per active user per day", "=" + ASSUMP["tok_user_day"], "60,000 (iternal.ai ranges).", NUM, GREEN)
w4 = wrow("W4", "Tokens per day", f"={w1}*{w2}*{w3}", "W1 x W2 x W3.", NUM)
w5 = wrow("W5", "Average tokens per second over a 10-hour working day", f"={w4}/(10*3600)", "W4 / (10 x 3,600).", NUM)
w6 = wrow("W6", "Peak tokens per second", f"={w5}*{ASSUMP['peak_factor']}", "W5 x peak factor (3x).", NUM)
w7 = wrow("W7", "Tokens per second one Blackwell GPU serves (70B-class model)", "=" + ASSUMP["tps_gpu_70b"], "2,500 tokens/s, conservative reading of NVIDIA's $0.02 per million tokens claim.", NUM, GREEN)
w8 = wrow("W8", "GPUs needed at peak, at target utilisation", f"={w6}/{w7}/{ASSUMP['util_inf']}", "W6 / W7 / utilisation (60%).", NUM, BLACK, True, KEYFILL)
w9 = wrow("W9", "Share of the assumed G-SIB fleet", f"={w8}/{ASSUMP['gpu_gsib']}", "W8 / GPUs per G-SIB adopter (6,000). A bank-wide assistant for 230,000 people fits in a few hundred GPUs. Inference for staff is not what fills the fleet.", PCT)
w10 = wrow("W10", "Annual tokens (250 working days), billions", f"={w4}*250/1e9", "W4 x 250 / 1e9.", NUM1)
w11 = wrow("W11", "Annual inference cost at NVIDIA's Blackwell $/M tokens ($M)", f"={w10}*1000*{ASSUMP['tok_cost_bw']}/1e6", "W10 (billions) x 1,000 x $ per million tokens / 1e6. Marginal compute cost of serving the bank at $0.02/M tokens.", USD2)
w12 = wrow("W12", "Same tokens at frontier API list price ($M)", f"={w10}*1000*{ASSUMP['api_price_frontier']}/1e6", "W10 x 1,000 x API price / 1e6. The gap between W11 and W12 is the prize for banks that own inference on open models.", USD1)

wsec("2. Real-time fraud scoring at a payment network (Visa scale)")
f1 = wrow("F1", "Transactions per year", "=" + ASSUMP["visa_txn_yr"], "269 billion (Visa).", NUM, GREEN)
f2 = wrow("F2", "Average transactions per second", f"={f1}/{ASSUMP['hours_yr']}/3600", "F1 / 8,760 / 3,600.", NUM)
f3 = wrow("F3", "Peak capacity transactions per second", "=" + ASSUMP["visa_tps_peak"], "76,000 TPS (Bloomberg).", NUM, GREEN)
f4 = wrow("F4", "FLOPs per scored transaction", "=" + ASSUMP["flop_txn"], "2 GFLOPs (1B-parameter model forward pass). Assumption.", NUM, GREEN)
f5 = wrow("F5", "Peak compute needed, PFLOPS", f"={f3}*{f4}/1e15", "F3 x F4 / 1e15.", '0.000')
f6 = wrow("F6", "Effective PFLOPS per GPU", "=" + ASSUMP["pflops_gpu_fp8"], "2 PFLOPS effective FP8 per Blackwell GPU.", NUM1, GREEN)
f7 = wrow("F7", "GPUs of raw compute to score every transaction at peak, with 2x redundancy", f"=2*{f5}/{f6}", "2 x F5 / F6. Scoring an entire global card network at peak needs less than one GPU of FLOPs. The real count is set by latency budgets, regions and redundancy, so a network runs dozens of inference GPUs, not thousands.", '0.00', BLACK, True, KEYFILL)
f8 = wrow("F8", "Training a transaction foundation model: tokens seen (transactions x 50 tokens each)", f"={ASSUMP['mc_txn_yr']}*50", "Mastercard 125B transactions x ~50 tokens per transaction record.", NUM)
f9 = wrow("F9", "Training FLOPs for the foundation model (6 x params x tokens)", f"=6*{ASSUMP['tfm_params']}*{f8}", "Standard 6ND estimate: 6 x parameters (10B) x tokens.", '0.00E+00')
f10 = wrow("F10", "GPU-hours at effective 2 PFLOPS", f"={f9}/({f6}*1e15)/3600", "F9 / (F6 x 1e15) / 3,600.", NUM)
f11 = wrow("F11", "GPUs to train in 30 days", f"={f10}/(30*24)", "F10 / 720 hours. One training run of a 10B transaction model on a year of data takes this many GPUs for a month; retraining monthly across several models and regions, plus research, is what justifies a 1,000-3,000 GPU payments fleet.", NUM, BLACK, True, KEYFILL)

wsec("3. Quant research cluster (XTX scale)")
q1 = wrow("Q1", "XTX Markets research GPUs", "=" + ASSUMP["xtx_gpus"], "Over 25,000 (XTX homepage).", NUM, GREEN)
q2 = wrow("Q2", "Ratio to the G-SIB employee assistant need (W8)", f"={q1}/{w8}", "Q1 / W8. One trading firm runs two orders of magnitude more GPUs than a 230,000-person bank's assistant needs.", MULT)
q3 = wrow("Q3", "Implied replacement value at $42k per installed GPU ($M)", f"={q1}*{ASSUMP['price_gpu_sys']}/1e6", "Q1 x price per installed GPU / 1e6. XTX's actual fleet is older A100/V100, bought cheaper; this is what refreshing it costs.", USD1)
q4 = wrow("Q4", "XTX disclosed data center capex, Finland ($M)", f"={ASSUMP['xtx_capex']}*1000", "More than EUR 1B (Bloomberg) for five data centers, buildings and power, not GPUs.", USD1, GREEN)
q5 = wrow("Q5", "HRT research iteration speedup Blackwell vs Hopper", "=" + ASSUMP["hrt_speedup"], "1.6x (NVIDIA GTC 2026). A quant firm's buying trigger is research velocity, not tokens.", MULT, GREEN)

wsec("Takeaway")
wrow("", "GPUs for the three builds: assistant (W8), fraud scoring (F7), one quant firm (Q1)", f'=TEXT({w8},"#,##0")&" / "&TEXT({f7},"0.00")&" / "&TEXT({q1},"#,##0")', "Text of W8 / F7 / Q1. Fleets in finance are driven by research and training, and inference count is set by latency, residency and redundancy rather than FLOPs.", '@', BLACK, True, KEYFILL)
W.freeze_panes = "A5"

# ---------------------------------------------------------------- TCO
C = wb.create_sheet("TCO")
title(C, "Five-year TCO: one 8-GPU Blackwell node on-premises vs the same GPUs rented in the cloud",
      "On-prem cost is largely fixed; cloud cost scales with hours used. The breakeven utilisation tells a bank when owning pays.")
header(C, 4, [1, 2, 3, 4], ["Line", "Item", "Value", "Formula (in words)"])
widths(C, [6, 70, 18, 110])
cr = 5
def crow(n, label, formula, words, fmt=USD, font=BLACK, bold=False, fill=None):
    global cr
    C.cell(row=cr, column=1, value=n).font = NOTE
    C.cell(row=cr, column=2, value=label).font = BOLD if bold else BLACK
    c = C.cell(row=cr, column=3, value=formula)
    c.font = Font(name=F, bold=bold, size=10, color=font.color)
    c.number_format = fmt
    w = C.cell(row=cr, column=4, value=words)
    w.font = NOTE
    w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 5):
            C.cell(row=cr, column=col).fill = fill
    cr += 1
    return f"C{cr-1}"
def csec(name):
    global cr
    C.cell(row=cr, column=2, value=name).font = BOLD
    for col in range(1, 5):
        C.cell(row=cr, column=col).fill = SECFILL
    cr += 1

csec("On-premises, one DGX B200 class node (8 GPUs), five years, $")
t1 = crow("T1", "System price", "=" + ASSUMP["node_price"], "$420,000, see Assumptions.", USD, GREEN)
t2 = crow("T2", "Networking, storage, installation", f"={t1}*{ASSUMP['node_netstor']}", "T1 x 25%.", USD)
t3 = crow("T3", "Capital cost", f"={t1}+{t2}", "T1 + T2.", USD)
t4 = crow("T4", "Average IT power draw, kW", f"={ASSUMP['node_kw']}*{ASSUMP['node_load']}", "14.3 kW max x 75% average load.", NUM1)
t5 = crow("T5", "Facility power incl. cooling, kW", f"={t4}*{ASSUMP['pue']}", "T4 x PUE (1.3).", NUM1)
t6 = crow("T6", "Electricity per year", f"={t5}*{ASSUMP['hours_yr']}*{ASSUMP['elec_price']}", "T5 x 8,760 hours x $/kWh (9.77c industrial, EIA Jul 2026).", USD)
t7 = crow("T7", "Colocation per year (billed on max IT kW)", f"={ASSUMP['node_kw']}*{ASSUMP['colo_kw_mo']}*12", "Max kW (14.3) x $204 per kW per month x 12. Colo contracts bill on committed power, not average draw.", USD)
t8 = crow("T8", "Software and support per year", f"=8*{ASSUMP['sw_gpu_yr']}", "8 GPUs x $4,500 per GPU per year (AI Enterprise).", USD)
t9 = crow("T9", "Operations staff per year", "=" + ASSUMP["staff_node_yr"], "$15,000 allocated per node.", USD, GREEN)
t10 = crow("T10", "Operating cost per year", f"={t6}+{t7}+{t8}+{t9}", "T6 + T7 + T8 + T9.", USD)
t11 = crow("T11", "Residual value recovered at end of year 5", f"={t1}*{ASSUMP['resid_val']}", "T1 x 10%.", USD)
t12 = crow("T12", "Five-year TCO, on-premises", f"={t3}+{t10}*{ASSUMP['yrs']}-{t11}", "T3 + T10 x 5 - T11.", USD, BLACK, True, KEYFILL)
t13 = crow("T13", "Share of five-year TCO that is the hardware", f"={t3}/{t12}", "T3 / T12.", PCT)
t14 = crow("T14", "Share that is power and facility", f"=({t6}+{t7})*{ASSUMP['yrs']}/{t12}", "(T6 + T7) x 5 / T12.", PCT)
t15 = crow("T15", "GPU-hours available over five years", f"=8*{ASSUMP['hours_yr']}*{ASSUMP['yrs']}", "8 x 8,760 x 5.", NUM)
t16 = crow("T16", "GPU-hours actually used at owned utilisation", f"={t15}*{ASSUMP['util_own']}", "T15 x 65%.", NUM)
t17 = crow("T17", "Effective cost per used GPU-hour, on-premises", f"={t12}/{t16}", "T12 / T16. Compare with cloud prices below.", USD2, BLACK, True, KEYFILL)
t18 = crow("T18", "Effective cost per available GPU-hour (100% utilisation)", f"={t12}/{t15}", "T12 / T15. The floor an owner can reach.", USD2)

csec("Cloud, same used GPU-hours, five years, $")
u1 = crow("U1", "Cloud cost at on-demand B200 price", f"={t16}*{ASSUMP['cloud_b200_hr']}", "T16 x $6.99 per GPU-hour.", USD, BLACK, True)
u2 = crow("U2", "Cloud cost at reserved B200 price", f"={t16}*{ASSUMP['cloud_b200_res']}", "T16 x $4.75 per GPU-hour.", USD, BLACK, True)
u3 = crow("U3", "On-prem saving vs on-demand cloud", f"={u1}-{t12}", "U1 - T12.", USD)
u4 = crow("U4", "On-prem saving vs reserved cloud", f"={u2}-{t12}", "U2 - T12. Negative means renting is cheaper at this utilisation.", USD)
u5 = crow("U5", "On-prem TCO as share of on-demand cloud", f"={t12}/{u1}", "T12 / U1.", PCT)
u6 = crow("U6", "Breakeven utilisation vs on-demand cloud", f"={t12}/({t15}*{ASSUMP['cloud_b200_hr']})", "T12 / (T15 x on-demand price). Above this share of hours busy, owning is cheaper than renting on demand.", PCT, BLACK, True, KEYFILL)
u7 = crow("U7", "Breakeven utilisation vs reserved cloud", f"={t12}/({t15}*{ASSUMP['cloud_b200_res']})", "T12 / (T15 x reserved price).", PCT, BLACK, True, KEYFILL)
u8 = crow("U8", "Same five-year used hours on H100 at the one-year rental price", f"={t16}*{ASSUMP['cloud_h100_1y']}", "T16 x $2.35. Note a B200 delivers several times the tokens of an H100 per hour, so this is not like for like; it shows why older-generation rental undercuts new-generation purchase for light workloads.", USD)

csec("Sensitivity: on-prem effective $/GPU-hour at different utilisation")
for u in (0.3, 0.45, 0.6, 0.75, 0.9):
    crow("", f"Utilisation {int(u*100)}%", f"={t12}/({t15}*{u})", "T12 / (T15 x utilisation).", USD2)

csec("Scale: a 64-node SuperPOD-class system (512 GPUs)")
s1 = crow("S1", "Capital cost, 64 nodes", f"=64*{t3}", "64 x T3.", USD)
s2 = crow("S2", "Facility power, kW", f"=64*{t5}", "64 x T5.", NUM)
s3 = crow("S3", "Operating cost per year, 64 nodes", f"=64*{t10}", "64 x T10. Networking fabric cost scales faster than linearly in reality; treat as a floor.", USD)
s4 = crow("S4", "Five-year TCO, 64 nodes ($M)", f"=64*{t12}/1e6", "64 x T12 / 1e6.", USD1, BLACK, True)
s5 = crow("S5", "GB200 NVL72 rack alternative: price per GPU vs DGX node", f"={ASSUMP['nvl72_price']}*1e6/72/({t3}/8)", "(NVL72 rack $2.5M / 72 GPUs) / (node capital / 8 GPUs). Below 1.0 means the rack is cheaper per GPU before its NVLink and liquid cooling advantages.", MULT)
C.freeze_panes = "A5"

# ---------------------------------------------------------------- Value
V = wb.create_sheet("Value")
title(V, "Value: what the compute buys, use case by use case",
      "Each block states the value evidence, the compute it plausibly runs on (from TCO and Workloads), and the ratio. Ratios above 10x mean compute price is not the constraint.")
header(V, 4, [1, 2, 3, 4], ["Line", "Item", "Value", "Formula (in words)"])
widths(V, [6, 70, 18, 110])
vr = 5
def vrow(n, label, formula, words, fmt=USD1, font=BLACK, bold=False, fill=None):
    global vr
    V.cell(row=vr, column=1, value=n).font = NOTE
    V.cell(row=vr, column=2, value=label).font = BOLD if bold else BLACK
    c = V.cell(row=vr, column=3, value=formula)
    c.font = Font(name=F, bold=bold, size=10, color=font.color)
    c.number_format = fmt
    w = V.cell(row=vr, column=4, value=words)
    w.font = NOTE
    w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 5):
            V.cell(row=vr, column=col).fill = fill
    vr += 1
    return f"C{vr-1}"
def vsec(name):
    global vr
    V.cell(row=vr, column=2, value=name).font = BOLD
    for col in range(1, 5):
        V.cell(row=vr, column=col).fill = SECFILL
    vr += 1

vsec("A. Fraud prevention at a payment network (Visa)")
va1 = vrow("VA1", "Fraud prevented per year ($B)", "=" + ASSUMP["visa_fraud"], "$28B (Bloomberg, 2024).", USD1, GREEN)
va2 = vrow("VA2", "AI and data infrastructure spend per year ($B)", f"={ASSUMP['visa_ai_spend']}/10", "$3B over ten years / 10.", USD2)
va3 = vrow("VA3", "Value to spend ratio", f"={va1}/{va2}", "VA1 / VA2. Every dollar of AI and data infrastructure protected roughly this many dollars.", MULT, BLACK, True, KEYFILL)
va4 = vrow("VA4", "Fraud prevented per transaction ($)", f"={va1}*1e9/{ASSUMP['visa_txn_yr']}", "VA1 x 1e9 / transactions per year.", '$0.000')
va5 = vrow("VA5", "GPUs needed to score at peak (from Workloads F7)", "=Workloads!C" + f7[1:], "Pulled from Workloads F7.", '0.0', GREEN)
va6 = vrow("VA6", "Five-year TCO of a 2,000-GPU payments fleet ($M)", f"={ASSUMP['gpu_payments']}/8*TCO!{t12}/1e6", "GPUs per payments adopter / 8 per node x node five-year TCO / 1e6. The fleet the bottom-up model assumes for a payments adopter.", USD1, GREEN)
va7 = vrow("VA7", "Annualised fleet cost as share of fraud prevented", f"=({va6}/5)/({va1}*1000)", "(VA6 / 5) / (VA1 x 1,000). Compute is a rounding error against the fraud it stops.", '0.00%', BLACK, True, KEYFILL)
va8 = vrow("VA8", "Incremental: 20% precision gain (Revolut) applied to the average organisation's $60M fraud loss ($M)", f"={ASSUMP['fraud_org_loss']}*{ASSUMP['revolut_prec']}", "$60M x 20%. What a transaction foundation model is worth to one mid-size issuer or acquirer per year, before false-positive savings.", USD1)
va9 = vrow("VA9", "Five-year TCO of a 64-GPU regional-bank pod ($M)", f"=8*TCO!{t12}/1e6", "8 nodes x node TCO / 1e6. Matches the regional bank fleet assumption.", USD1, GREEN)
va10 = vrow("VA10", "Payback of that pod on VA8 alone (years)", f"={va9}/{va8}", "VA9 / VA8.", '0.0', BLACK, True, KEYFILL)

vsec("B. Enterprise productivity at a G-SIB (JPMorgan)")
vb1 = vrow("VB1", "Reported AI savings per year ($B)", "=" + ASSUMP["jpm_savings"], "$2B (Forbes, Jul 2026).", USD1, GREEN)
vb2 = vrow("VB2", "Reported AI investment per year ($B)", "=" + ASSUMP["jpm_ai_spend"], "$2B (Olakai citing JPMorgan).", USD1, GREEN)
vb3 = vrow("VB3", "Savings to investment ratio", f"={vb1}/{vb2}", "VB1 / VB2. Roughly break-even on the bank's own numbers, which is why budgets keep rising rather than exploding.", MULT, BLACK, True, KEYFILL)
vb4 = vrow("VB4", "Theoretical productivity value: users x loaded cost x time share x efficiency gain ($B)", f"={ASSUMP['jpm_users']}*{ASSUMP['loaded_cost']}*{ASSUMP['time_ai_share']}*{ASSUMP['jpm_eff']}/1e9", "230,000 x $150,000 x 25% x 35% / 1e9. The upper bound if self-reported gains convert to cost.", USD1)
vb5 = vrow("VB5", "Realised share of the theoretical value", f"={vb1}/{vb4}", "VB1 / VB4.", PCT)
vb6 = vrow("VB6", "Five-year TCO of the GPUs the assistant needs (Workloads W8) ($M)", f"=Workloads!{w8}/8*TCO!{t12}/1e6", "W8 GPUs / 8 x node TCO / 1e6.", USD1, GREEN)
vb7 = vrow("VB7", "Annualised assistant compute as share of reported savings", f"=({vb6}/5)/({vb1}*1000)", "(VB6 / 5) / (VB1 x 1,000). The GPU bill for the assistant is small; integration, data and change cost dominate the $2B investment.", PCT, BLACK, True, KEYFILL)
vb8 = vrow("VB8", "Technology budget ($B)", "=" + ASSUMP["jpm_tech"], "$19.8B (JPMorganChase).", USD1, GREEN)
vb9 = vrow("VB9", "AI investment as share of technology budget", f"={vb2}/{vb8}", "VB2 / VB8.", PCT)

vsec("C. Risk engine consolidation at a dealer bank (Murex XVA on GPU)")
vc1 = vrow("VC1", "CPU servers in the XVA grid", "=" + ASSUMP["risk_grid_servers"], "5,000 (assumption).", NUM, GREEN)
vc2 = vrow("VC2", "Performance gain per GPU server vs CPU server", "=" + ASSUMP["murex_perf"], "7x (NVIDIA partners page, Murex).", MULT, GREEN)
vc3 = vrow("VC3", "GPU nodes to replace the grid at equal throughput", f"={vc1}/({vc2}*8)", "VC1 / (VC2 x 8 GPUs per node). Assumes one GPU ~ one CPU server x 7.", NUM)
vc4 = vrow("VC4", "CPU grid power, kW", f"={vc1}*{ASSUMP['cpu_server_kw']}", "VC1 x 1 kW.", NUM)
vc5 = vrow("VC5", "GPU replacement power, kW", f"={vc3}*{ASSUMP['node_kw']}*{ASSUMP['node_load']}", "VC3 x 14.3 kW x 75%.", NUM)
vc6 = vrow("VC6", "Power reduction", f"=1-{vc5}/{vc4}", "1 - VC5 / VC4. Murex quotes 4x less energy; this build lands in the same region.", PCT)
vc7 = vrow("VC7", "Annual power and colo saving ($M)", f"=(({vc4}-{vc5})*{ASSUMP['pue']}*{ASSUMP['hours_yr']}*{ASSUMP['elec_price']}+({vc4}-{vc5})*{ASSUMP['colo_kw_mo']}*12)/1e6", "((kW saved x PUE x 8,760 x $/kWh) + (kW saved x colo $/kW/mo x 12)) / 1e6.", USD1)
vc8 = vrow("VC8", "CPU grid refresh avoided, five-year ($M)", f"={vc1}*{ASSUMP['cpu_server_price']}/1e6", "VC1 x $25,000 / 1e6.", USD1)
vc9 = vrow("VC9", "GPU capital required ($M)", f"={vc3}*TCO!{t3}/1e6", "VC3 x node capital cost / 1e6.", USD1, GREEN)
vc10 = vrow("VC10", "Five-year net saving before the value of intraday risk ($M)", f"={vc8}+{vc7}*5-{vc9}", "VC8 + VC7 x 5 - VC9. Riskfuel's 'once a day to on demand' is the real prize and is not priced here.", USD1, BLACK, True, KEYFILL)

vsec("D. Quant research (XTX, HRT)")
vd1 = vrow("VD1", "Citadel Securities net trading revenue 2025 ($B)", 12.2, "quantvero 2026 summary; HRT ~$12.3B. Secondary source, order of magnitude only. https://www.quantvero.com/algo-trading/best-quant-trading-firms/", USD1, BLUE)
vd2 = vrow("VD2", "Annualised cost of a 25,000-GPU research fleet at on-prem effective $/GPU-hour ($B)", f"=25000*{ASSUMP['hours_yr']}*{ASSUMP['util_own']}*TCO!{t17}/1e9", "25,000 x 8,760 x 65% x effective $/GPU-hour / 1e9.", USD2, GREEN)
vd3 = vrow("VD3", "Fleet cost as share of a top firm's trading revenue", f"={vd2}/{vd1}", "VD2 / VD1. A full XTX-scale fleet costs a top market maker single-digit percent of revenue. Research velocity (HRT: 1.6x) is worth far more than that.", PCT, BLACK, True, KEYFILL)

vsec("Summary: annual compute cost as share of annual value, by use case")
vrow("", "Payments fraud (VA7)", f"={va7}", "", '0.00%')
vrow("", "G-SIB assistant (VB7)", f"={vb7}", "", PCT)
vrow("", "Quant research fleet vs trading revenue (VD3)", f"={vd3}", "", PCT)
vrow("", "Reading", "", "In every case compute is a small share of the value created. The constraint on GPU demand in finance is model and data readiness, regulation and power, not GPU price. That argues for value-based packaging (outcome-linked NIM blueprints) over price cuts.", '@')
V.freeze_panes = "A5"


# ---------------------------------------------------------------- Configs
G = wb.create_sheet("Configs")
title(G, "System configurations financial institutions deploy: four reference builds evaluated side by side",
      "One column per configuration. Rows are the specification, the five-year cost, the unit economics and the workload fit. Every cost row reuses the TCO sheet's rates so the four columns are comparable.")
header(G, 4, [1, 2, 3, 4, 5, 6], ["Line", "Item", "Regulated pod", "Bank SuperPOD class", "Sovereign AI factory", "Research factory", ])
widths(G, [6, 58, 20, 20, 20, 20])
G.cell(row=5, column=2, value="Who deploys it").font = BOLD
for col, txt in zip((3,4,5,6), ["Regional bank, mid-size insurer, exchange", "G-SIB or top-100 bank first system (BNY 2024)", "G-SIB, payment network, sovereign cloud for banks", "Quant trading or market-making firm"]):
    c = G.cell(row=5, column=col, value=txt); c.font = NOTE; c.alignment = Alignment(wrap_text=True, vertical="top")
G.row_dimensions[5].height = 42
gr = 6
def grow(n, label, f3, f4, f5, f6, fmt=NUM, bold=False, fill=None, fonts=None):
    global gr
    G.cell(row=gr, column=1, value=n).font = NOTE
    G.cell(row=gr, column=2, value=label).font = BOLD if bold else BLACK
    for col, f in zip((3,4,5,6), (f3,f4,f5,f6)):
        c = G.cell(row=gr, column=col, value=f)
        fnt = (fonts[col-3] if fonts else BLACK)
        c.font = Font(name=F, bold=bold, size=10, color=fnt.color)
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right")
    if fill:
        for col in range(1, 7):
            G.cell(row=gr, column=col).fill = fill
    gr += 1
    return gr - 1
def gsec(name):
    global gr
    G.cell(row=gr, column=2, value=name).font = BOLD
    for col in range(1, 7):
        G.cell(row=gr, column=col).fill = SECFILL
    gr += 1

gsec("Specification")
rN = grow("K1", "Nodes or racks", 4, 64, 4, 250, NUM, fonts=[BLUE]*4)
rT = grow("K2", "Unit type (1 = 8-GPU air-cooled node, 2 = GB300 NVL72 liquid-cooled rack)", 1, 1, 2, 1, NUM, fonts=[BLUE]*4)
rG = grow("K3", "GPUs", f"=C{rN}*8", f"=D{rN}*8", f"=E{rN}*72", f"=F{rN}*8", NUM, True)
rM = grow("K4", "GPU memory, TB", f"=C{rG}*{ASSUMP['b200_mem']}/1000", f"=D{rG}*{ASSUMP['b200_mem']}/1000", f"=E{rG}*{ASSUMP['b300_mem']}/1000", f"=F{rG}*{ASSUMP['b200_mem']}/1000", NUM1)
rP = grow("K5", "FP8 PFLOPS (NVIDIA sparse figures)", f"=C{rN}*{ASSUMP['node_fp8']}", f"=D{rN}*{ASSUMP['node_fp8']}", f"=E{rN}*{ASSUMP['node_fp8']}*{ASSUMP['nvl72_fp8_mult']}", f"=F{rN}*{ASSUMP['node_fp8']}", NUM)
rI = grow("K6", "Scale-up domain, GPUs sharing memory over NVLink", 8, 8, 72, 8, NUM, fonts=[BLUE]*4)
rF = grow("K7", "Scale-out fabric", "Top-of-rack Ethernet", "InfiniBand (SuperPOD RA)", "Spectrum-X or InfiniBand", "Spectrum-X Ethernet", '@', fonts=[BLUE]*4)
rC = grow("K8", "Cooling", "Air", "Air", "Liquid", "Air (liquid for Rubin)", '@', fonts=[BLUE]*4)
rW = grow("K9", "IT power, kW (max)", f"=C{rN}*{ASSUMP['node_kw']}", f"=D{rN}*{ASSUMP['node_kw']}", f"=E{rN}*{ASSUMP['gb300_kw']}", f"=F{rN}*{ASSUMP['node_kw']}", NUM)
rWf = grow("K10", "Facility power incl. cooling, kW (average load x PUE)", f"=C{rW}*{ASSUMP['node_load']}*{ASSUMP['pue']}", f"=D{rW}*{ASSUMP['node_load']}*{ASSUMP['pue']}", f"=E{rW}*{ASSUMP['node_load']}*{ASSUMP['hrt_pue']}", f"=F{rW}*{ASSUMP['node_load']}*{ASSUMP['hrt_pue']}", NUM)
rL = grow("K11", "Where it sits", "Bank's own data center or local colo cage", "Colo campus in a money-center city", "Colo campus in jurisdiction, 10-40 MW contract", "Owned or leased DC sited for cheap power", '@', fonts=[BLUE]*4)
rS = grow("K12", "Software stack", "AI Enterprise + NIM blueprints + EDB/Cloudera", "AI Enterprise + Base Command + Run:ai", "AI Enterprise + NeMo + Nemotron", "Own stack (PyTorch, Ray, custom)", '@', fonts=[BLUE]*4)

gsec("Five-year cost, $")
rU = grow("K13", "Unit price", f"={ASSUMP['node_price']}", f"={ASSUMP['node_price']}", f"={ASSUMP['gb300_price']}*1e6", f"={ASSUMP['hgx_air_price']}", USD, fonts=[GREEN]*4)
rCap = grow("K14", "Capital incl. networking, storage, installation", f"=C{rN}*C{rU}*(1+{ASSUMP['node_netstor']})", f"=D{rN}*D{rU}*(1+{ASSUMP['node_netstor']})", f"=E{rN}*E{rU}*(1+{ASSUMP['rack_fabric_nvl']})", f"=F{rN}*F{rU}*(1+{ASSUMP['node_netstor']})", USD, True)
rEl = grow("K15", "Electricity per year", f"=C{rWf}*{ASSUMP['hours_yr']}*{ASSUMP['elec_price']}", f"=D{rWf}*{ASSUMP['hours_yr']}*{ASSUMP['elec_price']}", f"=E{rWf}*{ASSUMP['hours_yr']}*{ASSUMP['elec_price']}", f"=F{rWf}*{ASSUMP['hours_yr']}*{ASSUMP['elec_price']}*0.6", USD)
rCo = grow("K16", "Colocation per year (committed kW x rate; liquid premium for racks; research factory owns its building at ~half the colo rate)", f"=C{rW}*{ASSUMP['colo_kw_mo']}*12", f"=D{rW}*{ASSUMP['colo_kw_mo']}*12", f"=E{rW}*{ASSUMP['colo_kw_mo']}*(1+{ASSUMP['liquid_premium']})*12", f"=F{rW}*{ASSUMP['colo_kw_mo']}*0.5*12", USD)
rSw = grow("K17", "Software and support per year (research factory runs its own stack)", f"=C{rG}*{ASSUMP['sw_gpu_yr']}", f"=D{rG}*{ASSUMP['sw_gpu_yr']}", f"=E{rG}*{ASSUMP['sw_gpu_yr']}", f"=F{rG}*{ASSUMP['sw_gpu_yr']}*0.25", USD)
rSt = grow("K18", "Operations staff per year", f"=C{rN}*{ASSUMP['staff_node_yr']}", f"=D{rN}*{ASSUMP['staff_node_yr']}", f"=E{rN}*9*{ASSUMP['staff_node_yr']}", f"=F{rN}*{ASSUMP['staff_node_yr']}", USD)
rOp = grow("K19", "Operating cost per year", f"=C{rEl}+C{rCo}+C{rSw}+C{rSt}", f"=D{rEl}+D{rCo}+D{rSw}+D{rSt}", f"=E{rEl}+E{rCo}+E{rSw}+E{rSt}", f"=F{rEl}+F{rCo}+F{rSw}+F{rSt}", USD, True)
rTco = grow("K20", "Five-year TCO (capital + 5 x operating - 10% residual on units)", f"=C{rCap}+5*C{rOp}-C{rN}*C{rU}*{ASSUMP['resid_val']}", f"=D{rCap}+5*D{rOp}-D{rN}*D{rU}*{ASSUMP['resid_val']}", f"=E{rCap}+5*E{rOp}-E{rN}*E{rU}*{ASSUMP['resid_val']}", f"=F{rCap}+5*F{rOp}-F{rN}*F{rU}*{ASSUMP['resid_val']}", USD, True, KEYFILL)
rTcoM = grow("K21", "Five-year TCO, $M", f"=C{rTco}/1e6", f"=D{rTco}/1e6", f"=E{rTco}/1e6", f"=F{rTco}/1e6", USD1, True)

gsec("Unit economics")
rUt = grow("K22", "Utilisation assumed", 0.4, f"={ASSUMP['util_own']}", 0.75, 0.85, PCT, fonts=[BLUE, GREEN, BLUE, BLUE])
rHr = grow("K23", "Used GPU-hours over five years", f"=C{rG}*{ASSUMP['hours_yr']}*5*C{rUt}", f"=D{rG}*{ASSUMP['hours_yr']}*5*D{rUt}", f"=E{rG}*{ASSUMP['hours_yr']}*5*E{rUt}", f"=F{rG}*{ASSUMP['hours_yr']}*5*F{rUt}", NUM)
rPh = grow("K24", "Effective $ per used GPU-hour", f"=C{rTco}/C{rHr}", f"=D{rTco}/D{rHr}", f"=E{rTco}/E{rHr}", f"=F{rTco}/F{rHr}", USD2, True, KEYFILL)
rPg = grow("K25", "Five-year TCO per GPU", f"=C{rTco}/C{rG}", f"=D{rTco}/D{rG}", f"=E{rTco}/E{rG}", f"=F{rTco}/F{rG}", USD)
rPp = grow("K26", "Capital per FP8 PFLOPS", f"=C{rCap}/C{rP}", f"=D{rCap}/D{rP}", f"=E{rCap}/E{rP}", f"=F{rCap}/F{rP}", USD)
rPw = grow("K27", "Five-year power and facility as share of TCO", f"=5*(C{rEl}+C{rCo})/C{rTco}", f"=5*(D{rEl}+D{rCo})/D{rTco}", f"=5*(E{rEl}+E{rCo})/E{rTco}", f"=5*(F{rEl}+F{rCo})/F{rTco}", PCT)
rCl = grow("K28", "Same used hours rented on demand ($6.99)", f"=C{rHr}*{ASSUMP['cloud_b200_hr']}", f"=D{rHr}*{ASSUMP['cloud_b200_hr']}", f"=E{rHr}*{ASSUMP['cloud_b200_hr']}", f"=F{rHr}*{ASSUMP['cloud_b200_hr']}", USD)
rCr = grow("K29", "Owned TCO as share of on-demand rental", f"=C{rTco}/C{rCl}", f"=D{rTco}/D{rCl}", f"=E{rTco}/E{rCl}", f"=F{rTco}/F{rCl}", PCT, True)

gsec("Workload fit (GPUs the configuration can carry, from Workloads)")
rWa = grow("K30", "Employee assistants of 230,000-user size it can serve", f"=C{rG}/Workloads!C{w8[1:]}", f"=D{rG}/Workloads!C{w8[1:]}", f"=E{rG}/Workloads!C{w8[1:]}", f"=F{rG}/Workloads!C{w8[1:]}", '0.0')
rWt = grow("K31", "10B-parameter transaction-model training runs it can complete per month", f"=C{rG}/Workloads!C{f11[1:]}", f"=D{rG}/Workloads!C{f11[1:]}", f"=E{rG}/Workloads!C{f11[1:]}", f"=F{rG}/Workloads!C{f11[1:]}", '0.0')
rWx = grow("K32", "Share of an XTX-scale research fleet", f"=C{rG}/{ASSUMP['xtx_gpus']}", f"=D{rG}/{ASSUMP['xtx_gpus']}", f"=E{rG}/{ASSUMP['xtx_gpus']}", f"=F{rG}/{ASSUMP['xtx_gpus']}", PCT)
rFit = grow("K33", "Primary workload", "Document processing, fraud model, staff assistant (inference)", "Risk models, fraud training, department assistants", "Transaction and risk foundation models (training), bank-wide inference", "Research, backtesting, strategy training", '@', fonts=[BLUE]*4)

gsec("Value fit (annualised TCO against the value evidence in the Value sheet)")
rV1 = grow("K34", "Annualised TCO, $M", f"=C{rTcoM}/5", f"=D{rTcoM}/5", f"=E{rTcoM}/5", f"=F{rTcoM}/5", USD2)
rV2 = grow("K35", "Value anchor used ($M per year)", f"=Value!C{va8[1:]}", f"=Value!C{vc7[1:]}+Value!C{vc8[1:]}/5", f"=Value!C{va1[1:]}*1000*0.05", f"={ASSUMP['hrt_speedup']}", USD1, fonts=[GREEN]*4)
rV3 = grow("K36", "What the anchor is", "20% precision gain on a $60M fraud loss (VA8)", "XVA grid consolidation: power saving plus avoided CPU refresh (VC7, VC8)", "5% of Visa-scale fraud prevented ($28B) attributable to model refresh", "Research iteration speedup, x (HRT); value is strategy P&L, not priced", '@', fonts=[BLUE]*4)
rV4 = grow("K37", "Annual value to annual compute cost", f"=C{rV2}/C{rV1}", f"=D{rV2}/D{rV1}", f"=E{rV2}/E{rV1}", "n/a", MULT, True, KEYFILL)
G.cell(row=gr, column=2, value="Formulas: GPUs = units x GPUs per unit. Capital = units x unit price x (1 + fabric share). Operating = electricity + colocation + software + staff, each from the TCO rates. TCO = capital + 5 x operating - residual. $/GPU-hour = TCO / (GPUs x 8,760 x 5 x utilisation). Value ratio = value anchor / (TCO / 5).").font = NOTE
G.freeze_panes = "C6"

# ---------------------------------------------------------------- Pricing
P = wb.create_sheet("Pricing")
title(P, "Pricing: what a financial institution pays per GPU-hour and per million tokens, by route, and where the margin sits",
      "Compares buying routes on one axis ($ per GPU-hour) and then on the unit the business understands (per million tokens).")
header(P, 4, [1, 2, 3, 4], ["Line", "Item", "Value", "Formula (in words)"])
widths(P, [6, 70, 18, 110])
pr = 5
def prow(n, label, formula, words, fmt=USD2, font=BLACK, bold=False, fill=None):
    global pr
    P.cell(row=pr, column=1, value=n).font = NOTE
    P.cell(row=pr, column=2, value=label).font = BOLD if bold else BLACK
    c = P.cell(row=pr, column=3, value=formula)
    c.font = Font(name=F, bold=bold, size=10, color=font.color)
    c.number_format = fmt
    w = P.cell(row=pr, column=4, value=words)
    w.font = NOTE
    w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 5):
            P.cell(row=pr, column=col).fill = fill
    pr += 1
    return f"C{pr-1}"
def psec(name):
    global pr
    P.cell(row=pr, column=2, value=name).font = BOLD
    for col in range(1, 5):
        P.cell(row=pr, column=col).fill = SECFILL
    pr += 1

psec("Price per GPU-hour by route (Blackwell B200 class unless noted)")
p1 = prow("P1", "Own, effective at 65% utilisation (TCO T17)", f"=TCO!{t17}", "From TCO.", USD2, GREEN)
p2 = prow("P2", "Own, at 100% utilisation (TCO T18)", f"=TCO!{t18}", "From TCO.", USD2, GREEN)
p3 = prow("P3", "Cloud reserved", "=" + ASSUMP["cloud_b200_res"], "$4.75.", USD2, GREEN)
p4 = prow("P4", "Cloud on-demand", "=" + ASSUMP["cloud_b200_hr"], "$6.99 (range $5.99 to $16.11).", USD2, GREEN)
p5 = prow("P5", "H100 one-year rental (previous generation)", "=" + ASSUMP["cloud_h100_1y"], "$2.35 (NVIDIA, Mar 2026).", USD2, GREEN)
p6 = prow("P6", "H100 on-demand median", "=" + ASSUMP["cloud_h100_hr"], "$2.70 (NVIDIA, Jun 2026). Up 35% in eight months: GPU rental prices rose in 2026, which is unusual for a compute market.", USD2, GREEN)
p7 = prow("P7", "Premium of on-demand cloud over owning (P4 / P1)", f"={p4}/{p1}", "P4 / P1.", MULT, BLACK, True, KEYFILL)

psec("Price per million tokens by route (70B-class open model)")
m1 = prow("M1", "Tokens one GPU serves per hour", f"={ASSUMP['tps_gpu_70b']}*3600", "2,500 tokens/s x 3,600.", NUM)
m2 = prow("M2", "Own: $ per million tokens at 65% utilisation", f"={p1}/{m1}*1e6", "P1 / M1 x 1e6.", '$0.000', BLACK, True)
m3 = prow("M3", "Cloud reserved: $ per million tokens", f"={p3}/{m1}*1e6", "P3 / M1 x 1e6.", '$0.000', BLACK, True)
m4 = prow("M4", "Cloud on-demand: $ per million tokens", f"={p4}/{m1}*1e6", "P4 / M1 x 1e6.", '$0.000', BLACK, True)
m5 = prow("M5", "NVIDIA's stated Blackwell floor (GPT-OSS-120B, TensorRT-LLM)", "=" + ASSUMP["tok_cost_bw"], "$0.02 per million tokens. Our M2 is above it because we assume 65% utilisation, 2,500 tokens/s and a full TCO stack.", '$0.000', GREEN)
m6 = prow("M6", "Hosted open-model API", "=" + ASSUMP["api_price_open"], "$0.60 blended list.", '$0.000', GREEN)
m7 = prow("M7", "Frontier API", "=" + ASSUMP["api_price_frontier"], "$6.00 blended list.", '$0.000', GREEN)
m8 = prow("M8", "Frontier API premium over owned inference (M7 / M2)", f"={m7}/{m2}", "M7 / M2. This spread is the economic case for banks to run open models on owned GPUs for the bulk of internal traffic, and NVIDIA's open-source pitch (84% of FS respondents say open source matters).", MULT, BLACK, True, KEYFILL)

psec("Where the margin sits in the stack")
g1 = prow("G1", "B200 GPU street price", "=" + ASSUMP["b200_chip"], "$47,500 (Oct 2026).", USD, GREEN)
g2 = prow("G2", "NVIDIA gross margin", "=" + ASSUMP["nvda_gm"], "75.0% GAAP, Q2 FY27.", PCT, GREEN)
g3 = prow("G3", "Implied NVIDIA cost of goods per GPU", f"={g1}*(1-{g2})", "G1 x (1 - G2).", USD)
g4 = prow("G4", "GPU share of the 8-GPU node price", f"=8*{g1}/{ASSUMP['node_price']}", "8 x G1 / node price. The GPUs are most of the box; CPUs, memory and chassis are the remainder.", PCT)
g5 = prow("G5", "GPU share of node capital cost incl. networking and storage", f"=8*{g1}/TCO!{t3}", "8 x G1 / TCO T3.", PCT)
g6 = prow("G6", "NVIDIA gross profit per 8-GPU node (chips only)", f"=8*({g1}-{g3})", "8 x (G1 - G3).", USD)
g7 = prow("G7", "NVIDIA software (AI Enterprise) over five years per node", f"=8*{ASSUMP['sw_gpu_yr']}*5", "8 x $4,500 x 5.", USD)
g8 = prow("G8", "Software as share of five-year node TCO", f"={g7}/TCO!{t12}", "G7 / TCO T12. Software is the recurring line NVIDIA can grow in FS without shipping more silicon.", PCT, BLACK, True, KEYFILL)
g9 = prow("G9", "Cloud provider gross margin on on-demand B200 at 65% utilisation", f"=1-{p1}/{p4}", "1 - P1 / P4. If the cloud's cost base matches our on-prem TCO.", PCT)

psec("Financing changes who pays")
h1 = prow("H1", "Third-party capital NVIDIA's partner platforms target ($B)", "=" + ASSUMP["fin_500"], "$500B (Aug 2026).", USD, GREEN)
h2 = prow("H2", "NVIDIA residual-value support, share of an opportunity", "=" + ASSUMP["fin_resid"], "Up to 25%.", PCT, GREEN)
h3 = prow("H3", "Five-year node TCO if hardware is financed with 25% residual support instead of 10% ($)", f"=TCO!{t3}+TCO!{t10}*5-TCO!{t1}*{h2}", "T3 + T10 x 5 - T1 x 25%.", USD)
h4 = prow("H4", "Saving vs base TCO", f"=TCO!{t12}-{h3}", "T12 - H3.", USD)
h5 = prow("H5", "Reading", "", "Goldman, Apollo, BlackRock, Blackstone, Brookfield and KKR now finance NVIDIA AI factories. Financial services is NVIDIA's lender as well as its customer. For the FS vertical, the pricing conversation moves from capex to lease rate per GPU-hour, which is the unit banks already use for cloud.", '@')
P.freeze_panes = "A5"

# ---------------------------------------------------------------- Sources
S = wb.create_sheet("Sources")
title(S, "Sources", "Numbering matches the article at esthercs.dev/research/fs-gpu.")
header(S, 4, [1, 2, 3, 4], ["No.", "Source", "Date", "URL"])
widths(S, [6, 90, 16, 110])
sources = [
    ("NVIDIA, Survey Reveals the Financial Services Industry Is Doubling Down on AI Investment and Open Source (State of AI in FS 2026)", "22 Jan 2026", "https://blogs.nvidia.com/blog/ai-in-financial-services-survey-2026/"),
    ("Gartner, Forecast: Enterprise IT Spending for the Banking and Investment Services Market, Worldwide, 2023-2029, 4Q25", "4Q25", "https://www.gartner.com/en/documents/7431362"),
    ("Gartner, Worldwide IT Spending to Grow 14.2% in 2026, Totaling $6.37 Trillion", "27 Jul 2026", "https://www.gartner.com/en/newsroom/press-releases/2026-07-27-gartner-forecasts-worldwide-it-spending-to-grow-14-point-2-percent-in-2026-totaling-6-point-37-trillion"),
    ("Gartner, Worldwide AI-Optimized IaaS Spending to Grow 96% in 2026", "10 Aug 2026", "https://www.gartner.com/en/newsroom/press-releases/2026-08-10-gartner-forecasts-worldwide-artificial-intelligence-optimized-iaas-spending-to-grow-96-percent-in-2026"),
    ("Gartner, Worldwide AI Spending to Grow 49.5% in 2026", "16 Sep 2026", "https://www.gartner.com/en/newsroom/press-releases/2026-09-16-gartner-forecasts-worldwide-ai-spending-to-grow-49-point-5-percent-in-2026"),
    ("IDC, Worldwide AI and Generative AI Spending, Industry Outlook", "21 Aug 2024", "https://www.idc.com/resource-center/blog/idcs-worldwide-ai-and-generative-ai-spending-industry-outlook/"),
    ("Digital Applied, AI Spending Forecasts 2026: Gartner, IDC and Stanford compiled (IDC AI infrastructure figures)", "2026", "https://www.digitalapplied.com/blog/ai-spending-forecasts-2026-gartner-idc-stanford-compiled"),
    ("NVIDIA, Financial Results for Second Quarter Fiscal 2027", "Aug 2026", "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-second-quarter-fiscal-2027"),
    ("NVIDIA, Financial Results for First Quarter Fiscal 2027", "May 2026", "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-first-quarter-fiscal-2027"),
    ("Jensen Huang, NVIDIA AI Factory Compute Is Becoming an Investable Asset Class", "12 Aug 2026", "https://blogs.nvidia.com/blog/nvidia-ai-factory-compute/"),
    ("FSB, 2025 List of Global Systemically Important Banks", "27 Nov 2025", "https://www.fsb.org/2025/11/2025-list-of-global-systemically-important-banks-g-sibs/"),
    ("S&P Global Market Intelligence, The world's largest banks by assets, 2026", "29 Apr 2026", "https://www.spglobal.com/market-intelligence/en/news-insights/research/2026/04/the-worlds-largest-banks-by-assets-2026"),
    ("FDIC, Quarterly Banking Profile, Second Quarter 2026", "2026", "https://www.fdic.gov/news/press-releases/2026/fdic-insured-institutions-reported-return-assets-137-percent-and-net"),
    ("BNY Mellon / NVIDIA, First Global Bank to Deploy AI Supercomputer Powered by NVIDIA DGX SuperPOD With DGX H100", "Mar 2024", "https://www.bny.com/corporate/global/en/about-us/newsroom/company-news/bny-mellon-first-global-bank-to-deploy-ai-supercomputer-powered-by-nvidia-dgx-superpod-with-dgx-h100.html"),
    ("Deutsche Bank, Deutsche Bank Partners with NVIDIA to embed AI into Financial Services", "7 Dec 2022", "https://www.db.com/news/detail/20221207-deutsche-bank-partners-with-nvidia-to-embed-ai-into-financial-services?language_id=1"),
    ("NVIDIA GTC 2026 live updates: Jump Trading, Hudson River Trading, Mastercard, Revolut, Adyen", "17-20 Mar 2026", "https://resources.nvidia.com/en-us-financial-services-industry/hudson-river"),
    ("XTX Markets homepage (over 25,000 GPUs)", "read 19 Oct 2026", "https://www.xtxmarkets.com/"),
    ("Fintech Blueprint, The secret AI supercomputers powering XTX Markets and DeepSeek", "18 Feb 2025", "https://lex.substack.com/p/ai-the-secret-ai-supercomputers-powering"),
    ("Bloomberg via Yahoo Finance, XTX to Build EUR 1 Billion Data Hub in Finland", "2024", "https://finance.yahoo.com/news/billionaire-alex-gerko-xtx-build-050000389.html"),
    ("WSJ, The Billionaire Math Geek Who Turned AI Into a Money-Printing Machine (XTX)", "23 Apr 2026", "https://www.wsj.com/finance/alex-gerko-xtx-markets-ai-d155626a"),
    ("Forbes / Bernard Marr, How JPMorgan Chase Is Building The AI-Powered Bank Of The Future", "1 Jul 2026", "https://www.forbes.com/sites/bernardmarr/2026/07/01/how-jpmorgan-chase-is-building-the-ai-powered-bank-of-the-future/"),
    ("JPMorganChase, technology page ($19.8B)", "2026", "https://www.jpmorganchase.com/about/technology"),
    ("Evident Insights, Banking Brief: AI remakes the org chart", "2026", "https://evidentinsights.com/bankingbrief/ai-remakes-the-org-chart/"),
    ("Mastercard, AI is helping banks save millions by transforming payment fraud prevention", "6 Feb 2026", "https://www.mastercard.com/us/en/news-and-trends/Insights/2026/ai-is-helping-banks-save-millions-by-transforming-payment-fraud-prevention.html"),
    ("Bloomberg Professional, Visa, Mastercard spur gen-AI adoption among payment companies", "4 Oct 2024", "https://www.bloomberg.com/professional/insights/artificial-intelligence/visa-mastercard-spur-gen-ai-adoption-among-payment-companies/"),
    ("Visa, AI solutions for fraud prevention and detection (269 billion transactions)", "read 19 Oct 2026", "https://corporate.visa.com/en/solutions/visa-protect/insights/ai-fraud-detection.html"),
    ("NVIDIA, Finance Solutions from Partners and ISVs", "3 Feb 2026", "https://www.nvidia.com/en-us/industries/finance/partners/"),
    ("NVIDIA, DGX B200 product page and specifications", "16 Oct 2026", "https://www.nvidia.com/en-us/data-center/dgx-b200/"),
    ("Tech-Insider, NVIDIA Blackwell GPU Pricing 2026: B200, B300 and DGX Costs", "Oct 2026", "https://tech-insider.org/nvidia-blackwell-gpu-pricing/"),
    ("Northflank, How much does an NVIDIA B200 GPU cost?", "2026", "https://northflank.com/blog/how-much-does-an-nvidia-b200-gpu-cost"),
    ("Spheron, H200 vs B200 vs GB200: Price and Memory", "16 Oct 2026", "https://www.spheron.network/blog/nvidia-h200-vs-b200-vs-gb200/"),
    ("CBRE, Global Data Center Trends 2026", "2026", "https://www.cbre.com/insights/reports/global-data-center-trends-2026"),
    ("Metro Colo Advisory, Data Center Cost 2026", "mid-2026", "https://metrocoloadvisory.com/data-center-cost/"),
    ("EIA, Electricity Monthly Update, End Use, July 2026", "Jul 2026", "https://www.eia.gov/electricity/monthly/update/end-use.php"),
    ("iternal.ai, Token Usage Guide 2026", "2026", "https://iternal.ai/token-usage-guide"),
    ("WEF, The AI Playbook for Financial Services", "Jun 2026", "https://reports.weforum.org/docs/WEF_The_AI_Playbook_for_Financial_Services_2026.pdf"),
    ("Cloud Security Alliance / Anjuna, The State of Cloud and AI for Financial Services 2026", "8 Jun 2026", "https://cloudsecurityalliance.org/artifacts/state-of-cloud-and-ai-for-financial-services-2026"),
    ("Reuters, Nvidia's China AI-chip share falls as Huawei closes gap (IDC)", "13 Oct 2026", "https://www.reuters.com/business/retail-consumer/nvidias-china-ai-chip-share-falls-huawei-closes-gap-report-finds-2026-10-13/"),
    ("CommandLinux, AI GPU Market Share NVIDIA vs AMD vs Intel: 2026 Statistics (IDC, Silicon Analysts)", "2026", "https://commandlinux.com/statistics/ai-gpu-market-share-nvidia-vs-amd-vs-intel/"),
    ("Digital Realty, The Frankfurt Advantage (IDC Tech Buyer Survey; 200 MW)", "16 Apr 2026", "https://www.digitalrealty.com/resources/blog/frankfurt-data-centers"),
    ("R. K. Thakur, NVIDIA's AI in Finance Report Just Confirmed What Most People Aren't Ready to Hear (839 respondents, 47% hybrid, 76% large firms)", "31 Mar 2026", "https://ninza7.medium.com/nvidias-ai-in-finance-report-just-confirmed-what-most-people-aren-t-ready-to-hear-2f26de01186d"),
    ("Air Street Capital, State of AI Report 2026", "Oct 2026", "https://www.stateof.ai/"),
    ("J.P. Morgan, Financing AI infrastructure and U.S. data centers ($697B hyperscaler capex)", "2026", "https://www.jpmorgan.com/insights/banking/capital-markets/financing-ai-infrastructure-data-centers"),
    ("Quantvero, 10 Best Quant Trading Firms (2026) (Citadel Securities, HRT revenue)", "2026", "https://www.quantvero.com/algo-trading/best-quant-trading-firms/"),
    ("Olakai, JP Morgan Spent $2B on AI. Here's What They Measured.", "2026", "https://olakai.ai/blog/jpmorgan-ai-measurement/"),
    ("NVIDIA Blog feed: LSEG, NatWest Nemotron Voice, Modulr", "Oct 2026", "https://blogs.nvidia.com/feed/"),
    ("AHEAD, NVIDIA 2026 NPN FSI Partner of the Year", "2026", "https://www.ahead.com/partner/nvidia/"),
]
sr = 5
for i, (s, d, u) in enumerate(sources, start=1):
    S.cell(row=sr, column=1, value=i).font = BLACK
    S.cell(row=sr, column=2, value=s).font = BLACK
    S.cell(row=sr, column=3, value=d).font = BLACK
    c = S.cell(row=sr, column=4, value=u)
    c.font = Font(name=F, size=9, color="0F6B63", underline="single")
    c.hyperlink = u
    sr += 1
S.freeze_panes = "A5"

wb.save(OUT)
print("saved", OUT, "assumption rows:", row - 5)
