"""Build the sovereign AI compute model workbook (companion to esthercs.dev/research/sovereign-ai).

Same conventions as the financial services model: every hardcoded number lives on
Assumptions with a source and read date, formulas reference those cells, and each
computed row says its formula in words.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUT = "/home/claude/ai-radar/research/sovereign-ai/sovereign-ai-model.xlsx"
F = "Arial"
BLUE = Font(name=F, color="0000FF", size=10); BLACK = Font(name=F, color="000000", size=10)
GREEN = Font(name=F, color="008000", size=10); BOLD = Font(name=F, bold=True, size=10)
H1 = Font(name=F, bold=True, size=14); H2 = Font(name=F, bold=True, size=11, color="FFFFFF")
NOTE = Font(name=F, italic=True, size=9, color="555555")
HEADFILL = PatternFill("solid", fgColor="132033"); KEYFILL = PatternFill("solid", fgColor="FFFF00"); SECFILL = PatternFill("solid", fgColor="E6F2F0")
USD = '$#,##0;($#,##0);-'; USD1 = '$#,##0.0;($#,##0.0);-'; USD2 = '$#,##0.00;($#,##0.00);-'
NUM = '#,##0;(#,##0);-'; NUM1 = '#,##0.0;(#,##0.0);-'; PCT = '0.0%;(0.0%);-'; MULT = '0.0x'

wb = Workbook()

def header(ws, row, cols, titles):
    for c, t in zip(cols, titles):
        cell = ws.cell(row=row, column=c, value=t); cell.font = H2; cell.fill = HEADFILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = H1
    if sub: ws["A2"] = sub; ws["A2"].font = NOTE
def widths(ws, w):
    for i, x in enumerate(w, start=1): ws.column_dimensions[get_column_letter(i)].width = x

# ---------------------------------------------------------------- README
ws = wb.active; ws.title = "README"
title(ws, "Sovereign AI compute: sizing, cost, pricing and value model", "Companion to esthercs.dev/research/sovereign-ai. Built 19 October 2026 by Esther Suebsubanant. Every input is dated and sourced on the Assumptions sheet.")
rows = [("How to read this workbook", ""), ("Blue text", "Hardcoded input, sourced or labelled as my assumption on Assumptions."), ("Black text", "Formula on the same sheet."), ("Green text", "Formula pulling from another sheet."), ("Yellow fill", "Scenario lever you are expected to change."), ("Formula column", "Each computed row states its formula in words."), ("", ""),
        ("Sheets", ""), ("Assumptions", "Every input with unit, value, read date and source URL."), ("TopDown", "Three paths to 2026 sovereign AI compute spend: NVIDIA's disclosed sovereign revenue grossed up, analyst market estimates, and public funding with private leverage."), ("BottomUp", "Programme by programme: announced capacity, timeline, megawatts energised in 2026, GPUs and capex, grouped by quadrant."), ("GigawattTCO", "What a gigawatt costs to own, by country: capex, power, delay, and the owner's effective $ per GPU-hour."), ("Configs", "Four reference builds, one per quadrant: compute exporter campus, industrial AI factory, EU gigafactory, subsidised national GPU pool."), ("Pricing", "GPU-hour prices by route and region, subsidy gaps, where the margin sits."), ("Value", "What the compute is worth: GDP uplift claims, revenue per compute dollar, economic value per dollar of investment."), ("Sources", "Numbered source list matching the article."), ("", ""), ("Units", "US dollars. Megawatts are IT load unless marked facility. Years are calendar years unless marked FY (NVIDIA fiscal year ends late January).")]
r = 4
for k, v in rows:
    ws.cell(row=r, column=1, value=k).font = BOLD if v == "" else BLACK
    c = ws.cell(row=r, column=2, value=v); c.font = BLACK; c.alignment = Alignment(wrap_text=True, vertical="top"); r += 1
widths(ws, [22, 120])

# ---------------------------------------------------------------- Assumptions
A = wb.create_sheet("Assumptions")
title(A, "Assumptions and inputs", "Blue = hardcoded. 'My assumption' = no public figure, my estimate, flagged for argument.")
header(A, 4, [1, 2, 3, 4, 5, 6], ["Key", "Input", "Value", "Unit", "Read / dated", "Source"])
widths(A, [8, 64, 16, 16, 14, 110])
ASSUMP = {}; row = 5
def section(name):
    global row
    A.cell(row=row, column=2, value=name).font = BOLD
    for col in range(1, 7): A.cell(row=row, column=col).fill = SECFILL
    row += 1
def inp(key, label, value, unit, dated, source, fmt=None, lever=False):
    global row
    A.cell(row=row, column=1, value=key).font = NOTE
    A.cell(row=row, column=2, value=label).font = BLACK
    c = A.cell(row=row, column=3, value=value); c.font = BLUE
    if fmt: c.number_format = fmt
    if lever: c.fill = KEYFILL
    A.cell(row=row, column=4, value=unit).font = BLACK
    A.cell(row=row, column=5, value=dated).font = BLACK
    A.cell(row=row, column=6, value=source).font = NOTE
    ASSUMP[key] = f"Assumptions!$C${row}"; row += 1

section("Top-down inputs")
inp("nvda_sov_fy26", "NVIDIA sovereign AI revenue, FY2026 (ended 25 Jan 2026)", 30, "$B", "25 Feb 2026", "NVIDIA Q4 FY26 earnings call as cited by State of AI Report 2026 and Fortune (15 Oct 2026, 'tripled to more than $30 billion'). https://fortune.com/2026/10/15/nvidias-european-spending-spree-is-about-much-more-than-europe/", USD1)
inp("nvda_sov_growth", "Growth in NVIDIA sovereign revenue, FY26 to calendar 2026", 0.40, "fraction", "My assumption", "It tripled in FY26. Pipeline (HUMAIN 600k GPUs, Korea 260k, Japan 27.5k Rubin, Nscale 58.6k) supports continued growth but power gates delivery. 40%. Lever.", PCT, True)
inp("nvda_share", "NVIDIA share of accelerator revenue in sovereign deals", 0.85, "fraction", "2026", "Higher than the ~80% global share because sovereign deals are mostly NVIDIA-branded AI factories; AMD (HUMAIN 1 GW), Qualcomm (200 MW) and Groq (19,000 LPUs) take the rest. Lever.", PCT, True)
inp("server_share", "Servers incl. GPUs as share of a data center's up-front capex", 0.64, "fraction", "Jun 2026", "Carnegie Endowment, The Compute Coalition: servers 64%, networking 11%, storage 5%, engineering and construction 11%, electrical 5%, thermal 2.5%, generators 2%. https://carnegieendowment.org/research/2026/06/the-compute-coalition-how-to-build-the-future-of-ai-in-the-free-world", PCT)
inp("mm_2025", "Sovereign AI market, 2025 (MarketsandMarkets)", 40.0, "$B", "19 Aug 2026", "MarketsandMarkets press release: $40.0B (2025) to $148.0B (2032), 20.6% CAGR. https://finance.yahoo.com/technology/ai/articles/sovereign-ai-market-reach-usd-151500186.html", USD1)
inp("mm_cagr", "Sovereign AI market CAGR 2026-2032 (MarketsandMarkets)", 0.206, "fraction", "19 Aug 2026", "Same release.", PCT)
inp("gvr_2026", "Sovereign AI market, 2026 (Grand View Research)", 48.9, "$B", "2026", "Grand View Research: $41.3B 2025, $48.9B 2026, $180.5B 2033, 20.5% CAGR. https://www.grandviewresearch.com/industry-analysis/sovereign-ai-market-report", USD1)
inp("gvr_hw", "Hardware share of the sovereign AI market (Grand View)", 0.48, "fraction", "2026", "Same source (48.0% in 2025).", PCT)
inp("prec_2026", "Sovereign AI infrastructure market, 2026 (Precedence Research)", 19.2, "$B", "2026", "Precedence Research: $15.0B 2025, $19.2B 2026, $177.1B 2035, 28% CAGR. https://www.precedenceresearch.com/sovereign-ai-infrastructure-market", USD1)
inp("sns_2025", "Sovereign AI infrastructure market, 2025 (SNS Insider)", 17.81, "$B", "2026", "SNS Insider: $17.81B 2025 to $91.99B 2035. https://www.snsinsider.com/reports/sovereign-ai-infrastructure-market-10918", USD2)
inp("ai_infra_26", "Worldwide AI infrastructure spending, 2026 (Gartner)", 1484.4, "$B", "16 Sep 2026", "Gartner, Worldwide AI Spending to Grow 49.5% in 2026. https://www.gartner.com/en/newsroom/press-releases/2026-09-16-gartner-forecasts-worldwide-ai-spending-to-grow-49-point-5-percent-in-2026", USD1)
inp("providers_2026", "AI providers' AI infrastructure spend, 2026 (Gartner)", 1000, "$B", "14 Oct 2026", "Gartner: AI providers expected to spend nearly $1 trillion on AI infrastructure in 2026 and $6 trillion more 2027-2030. https://www.gartner.com/en/newsroom/press-releases/2026-10-14-gartner-says-over-us-dollars-three-and-a-half-trillion-of-artificial-intelligence-infrastructure-spending-could-be-exposed-to-impairment-by-the-end-of-2030", USD)
inp("hyper_capex", "Four US hyperscalers' 2026 capex guidance", 733, "$B", "Oct 2026", "State of AI Report 2026 (~$733B); J.P. Morgan estimates $697B; Carnegie ~$670B. https://www.stateof.ai/", USD)

section("Public funding commitments (Path C)")
inp("eu_giga", "EU and national funding for AI Gigafactories", 10, "EUR B", "12 Aug 2026", "European Commission AI Factories page: up to EUR 10B EU and national funding, to unlock at least EUR 20B private; up to seven gigafactories of >100,000 processors. https://digital-strategy.ec.europa.eu/en/policies/ai-factories", NUM)
inp("eu_factories", "EuroHPC supercomputing and AI Factories budget, 2021-2027", 10, "EUR B", "12 Aug 2026", "Same page (EUR 10B over 2021-2027; 19 AI Factories selected).", NUM)
inp("eu_years", "Years over which the EU money is spent", 5, "years", "My assumption", "2026-2030 for gigafactories; AI Factories run 2021-2027.", NUM)
inp("eur_usd", "EUR to USD", 1.08, "USD per EUR", "Oct 2026", "Approximate 2026 rate. My figure.", '0.00')
inp("jp_budget", "Japan METI/NEDO/IPA AI and semiconductor budget, FY2026", 7.9, "$B", "May 2026", "Augur, Who Owns AI? (~JPY 1.23T). Includes Rapidus. https://whoownsai.org/who-owns-ai.pdf", USD1)
inp("jp_compute_share", "Share of Japan's budget that is compute rather than fabs", 0.4, "fraction", "My assumption", "Rapidus takes the majority. METI GPU cloud subsidies were JPY 72.5B; GENIAC adds compute. 40%. Lever.", PCT, True)
inp("jp_rubin_inv", "Japan Rubin consortium (Noetra) planned investment over five years", 6.7, "$B", "Jul 2026", "Spheron: ~JPY 1 trillion over five years behind 27,500 Rubin GPUs. https://www.spheron.network/blog/gpu-cloud-providers-japan-2026/", USD1)
inp("kr_gpus_gov", "Korea MSIT sovereign GPUs (National AI Computing Center and sovereign clouds)", 50000, "GPUs", "31 Oct 2025", "NVIDIA press release: up to 50,000 GPUs via MSIT, initial 13,000 Blackwell. https://investor.nvidia.com/news/press-release-details/2025/NVIDIA-South-Korea-Government-and-Industrial-Giants-Build-AI-Infrastructure-and-Ecosystem-to-Fuel-Korea-Innovation-Industries-and-Jobs/default.aspx", NUM)
inp("kr_years", "Years over which Korea's government GPUs deploy", 3, "years", "31 Oct 2025", "Same release ('over the next several years').", NUM)
inp("in_budget", "IndiaAI Mission public commitment over five years", 1.2, "$B", "19 Aug 2026", "MarketsandMarkets: exceeding $1B over five years. https://finance.yahoo.com/technology/ai/articles/sovereign-ai-market-reach-usd-151500186.html", USD1)
inp("uk_budget", "UK AI compute commitments (AI and quantum package, AIRR to 2030, Sovereign AI Unit)", 4.5, "GBP B", "May 2026", "Augur: GBP 2.5B AI+quantum package (Mar 2026), AIRR >GBP 1B by 2030, Sovereign AI Unit GBP 500M, GBP 80M procurement. https://whoownsai.org/who-owns-ai.pdf", NUM1)
inp("uk_years", "Years over which UK money is spent", 4, "years", "My assumption", "To 2030.", NUM)
inp("gbp_usd", "GBP to USD", 1.30, "USD per GBP", "Oct 2026", "Approximate. My figure.", '0.00')
inp("fr_bpi", "Bpifrance AI commitment by 2029", 10, "EUR B", "Feb 2025", "Augur / TechCrunch 7 Feb 2025. https://whoownsai.org/who-owns-ai.pdf", NUM)
inp("fr_years", "Years for Bpifrance commitment", 5, "years", "", "2025-2029.", NUM)
inp("my_budget", "Malaysia 2026 budget for AI incl. sovereign AI cloud", 1.41, "$B", "2026", "Introl: MYR 5.9B ($1.41B) for AI; Digital in Asia: MYR 2.1B (~$490M) for a sovereign AI cloud. https://introl.com/blog/malaysia-thailand-emerging-ai-data-center-hubs-southeast-asia", USD2)
inp("sa_nif", "Saudi National Infrastructure Fund financing framework for HUMAIN (250 MW)", 1.2, "$B", "2025", "Saudi Market Research Consulting summary of HUMAIN/Infra agreement (non-binding). https://saudimarketresearchconsulting.com/insights/articles/humain-saudi-arabia-a-bold-sovereign-ai-push-reshaping-infrastructure-and-markets", USD1)
inp("ca_budget", "Canada Sovereign AI Compute Strategy", 2.0, "CAD B", "2024-2025", "Canadian federal budget commitment, widely reported; order of magnitude. My figure.", NUM1)
inp("cad_usd", "CAD to USD", 0.73, "USD per CAD", "Oct 2026", "Approximate. My figure.", '0.00')
inp("ca_years", "Years for Canada's programme", 5, "years", "", "", NUM)
inp("leverage", "Private and sovereign-wealth capital mobilised per dollar of public budget", 2.0, "x", "My assumption", "The Commission targets 10x private leverage for gigafactories; Japan's Noetra is 44 private companies behind public subsidy; Gulf capital is sovereign wealth rather than budget. 2x is deliberately conservative. Lever.", MULT, True)

section("Cost of a gigawatt (Epoch AI, GB200 NVL72 build, US hyperscaler)")
inp("capex_gw", "Up-front capex for a 1 GW IT-load AI data center", 38, "$B", "19 May 2026", "Epoch AI, Total cost of ownership of a one-gigawatt AI data center. https://epoch.ai/data-insights/ai-datacenter-cost-breakdown", USD)
inp("opex_gw", "Annual operating cost for 1 GW", 0.9, "$B per year", "14 May 2026", "Same source ($0.9B/yr).", USD1)
inp("tco_gw_yr", "Annualised total cost of ownership for 1 GW", 8.5, "$B per year", "14 May 2026", "Same source ($8.5B/yr; servers $5B = 60%; network $1.4B; facility $1.2B; energy $590M; labour $140M; maintenance $120M).", USD1)
inp("energy_gw_yr", "Energy cost for 1 GW per year (US)", 0.59, "$B per year", "14 May 2026", "Same source.", USD2)
inp("server_life", "IT equipment lifespan assumed", 4, "years", "14 May 2026", "Epoch: 3 years gives $12B/yr, 7 years gives $7B/yr; the base case implies about 4. With facility over 10 years this reproduces Epoch's $8.5B/yr within 2%.", NUM)
inp("facility_life", "Facility, network and electrical lifespan", 10, "years", "My assumption", "Buildings last longer, network and electrical gear less. 10 years blended.", NUM)
inp("hurdle", "Cost of capital applied to capex (Gartner hurdle rate)", 0.12, "fraction", "14 Oct 2026", "Gartner assumes a 12% return on invested capital as the hurdle for AI data center owners. https://www.gartner.com/en/newsroom/press-releases/2026-10-14-gartner-says-over-us-dollars-three-and-a-half-trillion-of-artificial-intelligence-infrastructure-spending-could-be-exposed-to-impairment-by-the-end-of-2030", PCT, True)
inp("rack_kw", "GB200 NVL72 rack power", 120, "kW", "2026", "Widely reported; GB300 racks 135-140 kW (vision2030.ai citing Tom's Hardware). https://vision2030.ai/analysis/humain-ai-infrastructure/", NUM)
inp("rack_gpus", "GPUs per NVL72 rack", 72, "GPUs", "", "", NUM)
inp("gb300_rack_kw", "GB300 NVL72 rack power", 137.5, "kW", "2026", "135-140 kW (vision2030.ai). Midpoint.", NUM1)
inp("util", "Utilisation of a sovereign AI factory", 0.70, "fraction", "Jun 2026", "Carnegie and BloombergNEF assume ~70% (BNEF: 69% average by 2035). Lever.", PCT, True)
inp("hours", "Hours per year", 8760, "hours", "", "", NUM)
inp("us_power", "US industrial electricity price", 0.0977, "$ per kWh", "Jul 2026", "EIA Electricity Monthly Update, end use, July 2026. https://www.eia.gov/electricity/monthly/update/end-use.php", '$0.0000')
inp("delay_1y", "Lifecycle value lost by a one-year delay to a US data center", 0.055, "fraction", "Jun 2026", "Carnegie: >5% of total value (>$500M on a 100 MW facility); 1.5-year delay -8.9%, 6-month -3.6%, 3-month -1.7%. https://carnegieendowment.org/research/2026/06/the-compute-coalition-how-to-build-the-future-of-ai-in-the-free-world", PCT)
inp("delay_cost_100mw", "Lifecycle value lost by a one-year delay to a 100 MW US data center", 500, "$M", "Jun 2026", "Same source (more than $500 million).", USD)
inp("gb200_hr", "GB200 lowest on-demand rental price", 10.50, "$ per GPU-hour", "May 2026", "Carnegie citing getdeploying.com (CoreWeave). Carnegie models $10.", USD2)

section("Country factors (Carnegie Compute Coalition lifecycle value index, US = 0)")
for key, name, val in [("cf_uae", "UAE", 0.035), ("cf_us", "United States", 0.0), ("cf_fi", "Finland", -0.01), ("cf_ca", "Canada", -0.036), ("cf_in", "India", -0.048), ("cf_fr", "France", -0.077), ("cf_kr", "South Korea", -0.125), ("cf_au", "Australia", -0.164), ("cf_uk", "United Kingdom", -0.187), ("cf_de", "Germany", -0.194)]:
    inp(key, f"Lifecycle value of a data center in {name} relative to the US, no delay", val, "fraction", "Jun 2026", "Carnegie Endowment, The Compute Coalition, country ranking figure. UK industrial power ~3x US; Germany has the longest timelines.", PCT)
inp("power_sa", "Industrial electricity price, Saudi Arabia", 0.05, "$ per kWh", "My assumption", "Regulated industrial tariff plus solar PPAs at record lows (Sakaka, Al Shuaibah, Sudair cited by vision2030.ai). 5 cents. Lever.", '$0.000', True)
inp("power_kr", "Industrial electricity price, South Korea", 0.13, "$ per kWh", "My assumption", "Carnegie ranks Korea -12.5% and notes prices 23% above the model band. 13 cents.", '$0.000')
inp("power_de", "Industrial electricity price, Germany", 0.18, "$ per kWh", "My assumption", "Carnegie: UK ~3x US; Germany similar order. 18 cents.", '$0.000')
inp("power_in", "Industrial electricity price, India", 0.09, "$ per kWh", "My assumption", "Carnegie ranks India -4.8%. 9 cents.", '$0.000')
inp("power_th", "Industrial electricity price, Thailand", 0.12, "$ per kWh", "My assumption", "Gas-heavy grid, 46 GW installed (Introl). 12 cents.", '$0.000')
inp("pue_gulf", "PUE in the Gulf (hot climate, liquid cooling)", 1.25, "PUE", "My assumption", "Liquid-cooled NVL72 helps; ambient hurts.", '0.00')
inp("pue_eu", "PUE in northern Europe / Korea", 1.2, "PUE", "My assumption", "HRT Lefdal 1.15; enterprise 1.2-1.5.", '0.00')
inp("delay_de", "Extra years to energise in Germany vs US", 1.5, "years", "Jun 2026", "Carnegie: Germany fares worst on timelines; Northumberland (UK) 720 MW takes 9 years vs Stargate UAE 1.4 GW in ~2.5 years (Epoch). 1.5 years. Lever.", NUM1, True)
inp("war_ins", "Gulf war-risk and country premium applied to data center capex per year", 0.025, "fraction of capex per year", "1 Apr 2026", "Semafor via Carnegie: Gulf war insurance rose 19x to $5M per $100M of coverage (5%) with reduced limits. Not all capex is insured at that rate. 2.5%. Lever.", PCT, True)

section("Pricing inputs ($ per GPU-hour, list, read Sep-Oct 2026)")
for key, label, val, src in [
    ("p_aws_eu", "AWS p5 H100, Stockholm", 7.36, "Qovery, Sep 2026 list. https://www.qovery.com/blog/european-gpu-cloud-alternatives-aws-azure-comparison"),
    ("p_azure_eu", "Azure ND H100 v5, West Europe", 15.98, "Qovery, Sep 2026."),
    ("p_scaleway", "Scaleway H100 PCIe, Paris (EUR 2.87)", 3.10, "Qovery, Sep 2026, converted at 1.08."),
    ("p_ovh", "OVHcloud H100, Gravelines", 2.99, "Qovery / Spheron, Sep-Oct 2026. https://www.spheron.network/blog/ovhcloud-gpu-pricing-2026-h100-cost-and-sovereign-ai/"),
    ("p_verda_b200", "Verda B200, Finland", 6.11, "Qovery, Sep 2026."),
    ("p_nebius_b200", "Nebius B200", 8.50, "computeprices.com, 14 Oct 2026. https://computeprices.com/providers/nebius"),
    ("p_nebius_h100", "Nebius H100 SXM", 4.50, "Same."),
    ("p_coreweave_h100", "CoreWeave HGX H100, UK/Norway/Sweden/Spain", 6.16, "Qovery, Sep 2026."),
    ("p_sakura", "Sakura Internet H100, Japan (JPY 990)", 6.07, "Spheron, Oct 2026. https://www.spheron.network/blog/gpu-cloud-providers-japan-2026/"),
    ("p_aws_jp", "AWS Japan H100", 12.30, "Spheron, Oct 2026."),
    ("p_india_sub", "IndiaAI subsidised high-end GPU rate (INR 150)", 1.70, "MeitY / IndiaAI; INR 65-150 vs INR 210-335 commercial. https://indiaai.gov.in/hub/indiaai-compute-capacity"),
    ("p_india_com", "Commercial H100 rate in India (INR 300 midpoint)", 3.40, "Fundingsstartup summary of IndiaAI comparison."),
    ("p_h100_median", "Global H100 on-demand median (NVIDIA index)", 2.70, "NVIDIA, Aug 2026 (June reading). https://blogs.nvidia.com/blog/nvidia-ai-factory-compute/"),
]:
    inp(key, label, val, "$ per GPU-hour", "Sep-Oct 2026", src, USD2)
inp("nvda_gm", "NVIDIA gross margin, Q2 FY27", 0.75, "fraction", "Aug 2026", "NVIDIA Q2 FY27 release. https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-second-quarter-fiscal-2027", PCT)
inp("gb300_chip", "GB300 chip value implied by US export approvals", 28571, "$ per chip", "Nov 2025", "vision2030.ai: 35,000 GB300-equivalent chips estimated at ~$1B, i.e. ~$28.6k each. Street B300 ~$53k (tech-insider). The gap is the allocation-price vs street-price spread.", USD)

section("Value inputs")
inp("pwc_gdp", "AI contribution to global GDP by 2030 (PwC)", 15.7, "$T", "2017", "PwC Sizing the Prize, via WEF. https://www.weforum.org/stories/geographies-in-depth/the-global-economy-will-be-14-bigger-in-2030-because-of-ai", USD1)
inp("pwc_pct", "GDP uplift by 2030 (PwC)", 0.14, "fraction", "2017", "Same.", PCT)
inp("idc_gdp", "Cumulative AI economic impact through 2030 (IDC)", 19.9, "$T", "2024", "IDC: $19.9T, 3.5% of global GDP in 2030. https://my.idc.com/getdoc.jsp?containerId=prUS52600524", USD1)
inp("th_value", "Economic value of Google's Bangkok cloud region over five years", 41, "$B", "21 Jan 2026", "Google Cloud / Public First: THB 1.4T. https://www.googlecloudpresscorner.com/2026-01-21-Google-Cloud-Launches-New-Cloud-Region-in-Thailand,-Bolstering-its-Commitment-to-Advancing-the-Countrys-AI-Driven-Digital-Economy", USD)
inp("th_invest", "Google's Thailand investment behind it", 1, "$B", "21 Jan 2026", "Same.", USD)
inp("anthropic_ratio", "Frontier lab revenue per dollar of compute", 1.70, "$ revenue per $ compute", "15 Oct 2026", "Fortune: Anthropic told investors $1.70 of revenue per $1 of compute; Q3 revenue $16.8B. https://fortune.com/2026/10/15/nvidias-european-spending-spree-is-about-much-more-than-europe/", USD2)
inp("gartner_rev_need", "Cumulative revenue AI data center owners need through 2030 (Gartner)", 8.5, "$T", "14 Oct 2026", "Gartner: to cover capex, financing and 12% ROIC. https://www.gartner.com/en/newsroom/press-releases/2026-10-14-gartner-says-over-us-dollars-three-and-a-half-trillion-of-artificial-intelligence-infrastructure-spending-could-be-exposed-to-impairment-by-the-end-of-2030", USD1)
inp("gartner_rev_have", "AI revenue attributable to AI data centers through 2030 (Gartner)", 3.6, "$T", "14 Oct 2026", "Same (no more than $3.6T).", USD1)
inp("gartner_impair", "AI infrastructure exposed to impairment by end-2030 (Gartner)", 3.5, "$T", "14 Oct 2026", "Same (>$3.5T; >$4T through 2032; 80% of neoclouds cease independent operation by 2030).", USD1)
inp("gdp_sa", "Saudi Arabia GDP", 1.1, "$T", "2025", "Approximate nominal GDP. My figure.", USD1)
inp("gdp_kr", "South Korea GDP", 1.8, "$T", "2025", "Approximate. My figure.", USD1)
inp("gdp_eu", "EU GDP", 19.0, "$T", "2025", "Approximate. My figure.", USD1)
inp("gdp_in", "India GDP", 4.0, "$T", "2025", "Approximate. My figure.", USD1)
inp("gdp_th", "Thailand GDP", 0.53, "$T", "2025", "Approximate. My figure.", USD2)
inp("ai_share_gdp", "Share of PwC's AI GDP uplift that depends on in-country compute (vs imported services)", 0.25, "fraction", "My assumption", "Most AI value is consumed as software and services that can run anywhere. Sovereign compute captures the part that must be local (regulated data, defence, latency, language). 25%. Lever.", PCT, True)
A.freeze_panes = "A5"

# ---------------------------------------------------------------- Programmes (bottom-up data)
# name, quadrant, country, announced MW (IT), target year, MW energised by end 2026 (my est), GPUs announced, public $B, source note
PROGS = [
    ("HUMAIN (PIF) AI factories incl. xAI 500 MW+, AMD 1 GW, stc 1 GW", "Compute exporters", "Saudi Arabia", 1900, 2030, 150, 600000, "1.9 GW by 2030, 6.6 GW by 2034; 50 MW pilot with 18,000 GB300 live 2026; AMD first 100 MW phase 2026. vision2030.ai; Saudi MRC."),
    ("Stargate UAE (G42, OpenAI, Oracle) within 5 GW campus", "Compute exporters", "UAE", 1000, 2029, 200, 0, "1 GW first phase ~$20B; 200 MW live 2026 (OpenAI, G42, Zawya)."),
    ("AWS Riyadh AI Zone (HUMAIN) and other Gulf (Qatar, QIA-Brookfield $20B JV)", "Compute exporters", "Gulf other", 400, 2029, 40, 150000, "AWS up to 150,000 accelerators in Riyadh, 100 MW with gigawatt ambition; QIA-Brookfield JV. My capacity estimate."),
    ("Korea sovereign GPUs: MSIT 50k, SK 50k, Hyundai 50k, NAVER 60k", "Industrial sovereigns", "South Korea", 500, 2028, 110, 260000, "NVIDIA release 31 Oct 2025: >260,000 GPUs. MW derived at ~1.9 kW per GPU installed."),
    ("Japan: Noetra 27,500 Rubin, SoftBank GB200 cloud, Sakura 22k GPUs, METI JPY 72.5B subsidies", "Industrial sovereigns", "Japan", 200, 2028, 40, 60000, "Spheron Oct 2026; GPU count sums Rubin 27.5k, Sakura ~21k, SoftBank est. 10k."),
    ("Taiwan NT$100B AI initiative and national AI supercomputer", "Industrial sovereigns", "Taiwan", 60, 2027, 20, 10000, "Augur (Jun 2025 initiative). ~10,000 Blackwell class. My MW estimate."),
    ("EU AI Gigafactories (up to 7 x 100,000 processors)", "Regulated builders", "EU", 1400, 2030, 0, 500000, "EC: tenders launched Aug 2026; EUR 10B public to unlock EUR 20B private. Nothing energised in 2026. 5 x 100k chips x ~2 kW."),
    ("EuroHPC AI Factories (19 selected, first 6 live Oct 2025)", "Regulated builders", "EU", 300, 2027, 120, 60000, "EC; Augur. ~15 MW average each. My estimate."),
    ("France: AI Campus JV 1.4 GW (ops 2028), Mistral Compute 40 MW (18,000 GB chips), UAE-France 1 GW", "Regulated builders", "France", 2400, 2030, 60, 18000, "Augur; Mistral Compute 40 MW live 2026 + early campus works."),
    ("Germany: Deutsche Telekom Industrial AI Cloud, 10,000 Blackwell GPUs, Munich", "Regulated builders", "Germany", 25, 2026, 25, 10000, "Telekom 4 Nov 2025; live Q1 2026; >1,000 DGX B200 at 14.3 kW."),
    ("UK: Nscale 58,640 GPUs incl. Loughton 50-90 MW (Q1 2027), Stargate UK, Isambard, AIRR", "Regulated builders", "UK", 200, 2028, 30, 58640, "Nscale 16 Sep 2025; OpenAI Stargate UK. Loughton delivers Q1 2027."),
    ("Nordics and other EU national (Denmark Gefion, Italy, Spain, Switzerland Alps, Norway)", "Regulated builders", "Europe other", 150, 2028, 50, 30000, "My estimate of national supercomputer and sovereign cloud builds."),
    ("IndiaAI Mission GPU pool (38,000-45,000 GPUs, subsidised at INR 150/hr)", "Subsidised tenants", "India", 90, 2027, 60, 45000, "MeitY; 45,000+ GPUs available Jun 2026. ~2 kW per GPU installed."),
    ("Malaysia: YTL 600 MW NVIDIA sovereign AI DC (Kulai), MYR 2.1B sovereign AI cloud", "Subsidised tenants", "Malaysia", 600, 2029, 100, 0, "Introl Dec 2025; Digital in Asia 2026."),
    ("Thailand: Gulf 2 GW AI DC plan, Siam AI, Google/Microsoft/AWS regions, BOI THB 521B approvals", "Subsidised tenants", "Thailand", 400, 2030, 30, 0, "LinkedIn (Gulf $4.3B, 2 GW); Digital in Asia. Sovereign portion is my estimate."),
    ("Indonesia, Vietnam, Singapore sovereign AI compute", "Subsidised tenants", "SEA other", 150, 2029, 40, 0, "My estimate."),
    ("Canada Sovereign AI Compute Strategy (CAD 2B) and Australia (Firmus $10B debt facility)", "Subsidised tenants", "Canada/Australia", 150, 2029, 30, 0, "Augur; my estimate."),
    ("Latin America (Brazil, Mexico, Chile) sovereign and national supercomputers", "Subsidised tenants", "LatAm", 80, 2029, 20, 0, "My estimate."),
]

B = wb.create_sheet("BottomUp")
title(B, "Bottom-up: sovereign AI compute programmes, capacity announced and energised in 2026",
      "MW are IT load. Energised-in-2026 is my estimate from each programme's disclosed phases. Capex per MW from Epoch ($38B/GW). GPUs per MW from NVL72 rack density. $M unless noted.")
header(B, 4, list(range(1, 13)), ["Programme", "Quadrant", "Country", "Announced MW (IT)", "Target year", "MW energised by end-2026", "GPUs announced", "GPUs implied in 2026 (MW x GPUs/MW)", "Capex energised 2026 ($M)", "Of which GPU systems ($M)", "Announced capex to target ($B)", "Source note"])
widths(B, [58, 18, 14, 12, 10, 14, 14, 16, 16, 16, 16, 70])
B.row_dimensions[4].height = 48
br = 5; first = br
for name, quad, ctry, mw, yr, mw26, gpus, note in PROGS:
    B.cell(row=br, column=1, value=name).font = BLACK
    B.cell(row=br, column=2, value=quad).font = BLACK
    B.cell(row=br, column=3, value=ctry).font = BLACK
    for col, v in ((4, mw), (5, yr), (6, mw26), (7, gpus)):
        c = B.cell(row=br, column=col, value=v); c.font = BLUE; c.number_format = NUM if col != 5 else '0'
    B.cell(row=br, column=6).fill = KEYFILL
    B.cell(row=br, column=8, value=f"=F{br}*1000/{ASSUMP['rack_kw']}*{ASSUMP['rack_gpus']}").number_format = NUM
    B.cell(row=br, column=9, value=f"=F{br}*{ASSUMP['capex_gw']}").number_format = USD
    B.cell(row=br, column=10, value=f"=I{br}*{ASSUMP['server_share']}").number_format = USD
    B.cell(row=br, column=11, value=f"=D{br}*{ASSUMP['capex_gw']}/1000").number_format = USD1
    B.cell(row=br, column=12, value=note).font = NOTE
    br += 1
last = br - 1
B.cell(row=br, column=1, value="Total").font = BOLD
for col in (4, 6, 7, 8, 9, 10, 11):
    L = get_column_letter(col)
    c = B.cell(row=br, column=col, value=f"=SUM({L}{first}:{L}{last})"); c.font = BOLD; c.fill = KEYFILL
    c.number_format = B.cell(row=last, column=col).number_format
tot = br; br += 2
B.cell(row=br, column=1, value="Formulas. GPUs implied 2026 = MW energised x 1,000 / rack kW x 72 GPUs per rack. Capex energised = MW x $38M per MW (Epoch, $38B per GW, all-in incl. facility). GPU systems = capex x 64% server share (Carnegie). Announced capex = announced MW x $38M.").font = NOTE
br += 2
B.cell(row=br, column=1, value="By quadrant, 2026").font = BOLD; br += 1
header(B, br, [1, 2, 3, 4, 5, 6], ["Quadrant", "MW energised 2026", "GPUs implied 2026", "Capex energised 2026 ($M)", "GPU systems 2026 ($M)", "Announced MW"]); br += 1
qstart = br
for q in ["Compute exporters", "Industrial sovereigns", "Regulated builders", "Subsidised tenants"]:
    B.cell(row=br, column=1, value=q).font = BLACK
    B.cell(row=br, column=2, value=f'=SUMIF($B${first}:$B${last},A{br},$F${first}:$F${last})').number_format = NUM
    B.cell(row=br, column=3, value=f'=SUMIF($B${first}:$B${last},A{br},$H${first}:$H${last})').number_format = NUM
    B.cell(row=br, column=4, value=f'=SUMIF($B${first}:$B${last},A{br},$I${first}:$I${last})').number_format = USD
    B.cell(row=br, column=5, value=f'=SUMIF($B${first}:$B${last},A{br},$J${first}:$J${last})').number_format = USD
    B.cell(row=br, column=6, value=f'=SUMIF($B${first}:$B${last},A{br},$D${first}:$D${last})').number_format = NUM
    br += 1
qend = br - 1
B.cell(row=br, column=1, value="Share of 2026 capex in the Gulf").font = BLACK
B.cell(row=br, column=4, value=f"=D{qstart}/I{tot}").number_format = PCT; br += 1
B.cell(row=br, column=1, value="Share of announced MW in Europe (Regulated builders)").font = BLACK
B.cell(row=br, column=6, value=f"=F{qstart+2}/D{tot}").number_format = PCT; br += 1
B.cell(row=br, column=1, value="Europe share of 2026 energised MW").font = BLACK
B.cell(row=br, column=2, value=f"=B{qstart+2}/F{tot}").number_format = PCT; br += 2
# Summary cells at fixed rows for cross-reference
B.cell(row=br, column=1, value="Summary 2026").font = BOLD; br += 1
SUMROW = br
B.cell(row=br, column=1, value="MW energised 2026 (GW)").font = BLACK
B.cell(row=br, column=3, value=f"=F{tot}/1000").number_format = '0.00'; br += 1
B.cell(row=br, column=1, value="Sovereign AI capex energised 2026, all-in ($B)").font = BOLD
c = B.cell(row=br, column=3, value=f"=I{tot}/1000"); c.number_format = USD1; c.font = BOLD; c.fill = KEYFILL; br += 1
B.cell(row=br, column=1, value="Of which GPU systems ($B)").font = BOLD
c = B.cell(row=br, column=3, value=f"=J{tot}/1000"); c.number_format = USD1; c.font = BOLD; c.fill = KEYFILL; br += 1
B.cell(row=br, column=1, value="GPUs implied by 2026 energisation").font = BLACK
B.cell(row=br, column=3, value=f"=H{tot}").number_format = NUM; br += 1
B.cell(row=br, column=1, value="Announced pipeline, all programmes (GW)").font = BLACK
B.cell(row=br, column=3, value=f"=D{tot}/1000").number_format = '0.0'; br += 1
B.cell(row=br, column=1, value="Announced pipeline capex ($B)").font = BLACK
B.cell(row=br, column=3, value=f"=K{tot}").number_format = USD; br += 1
B.cell(row=br, column=1, value="Years to deliver pipeline at 2026 pace").font = BLACK
B.cell(row=br, column=3, value=f"=D{tot}/F{tot}").number_format = '0.0'
B.cell(row=br, column=4, value="Announced MW / MW energised in 2026. Shows how far announcements run ahead of delivery.").font = NOTE; br += 1
B.freeze_panes = "B5"
BU = {"gw26": f"BottomUp!$C${SUMROW}", "capex26": f"BottomUp!$C${SUMROW+1}", "gpu26": f"BottomUp!$C${SUMROW+2}", "gpus26": f"BottomUp!$C${SUMROW+3}", "pipe_gw": f"BottomUp!$C${SUMROW+4}", "pipe_capex": f"BottomUp!$C${SUMROW+5}", "pipe_years": f"BottomUp!$C${SUMROW+6}"}

# ---------------------------------------------------------------- TopDown
T = wb.create_sheet("TopDown")
title(T, "Top-down: three paths to 2026 sovereign AI compute spend", "$B. Path A grosses up NVIDIA's disclosed sovereign revenue to all vendors and to the full data center. Path B uses analyst market estimates. Path C starts from public budgets and applies leverage.")
header(T, 4, [1, 2, 3, 4], ["Line", "Item", "Value", "Formula (in words) and source"])
widths(T, [6, 72, 18, 110])
tr = 5
def trow(n, label, formula, words, fmt=USD1, font=BLACK, bold=False, fill=None):
    global tr
    T.cell(row=tr, column=1, value=n).font = NOTE
    T.cell(row=tr, column=2, value=label).font = BOLD if bold else BLACK
    c = T.cell(row=tr, column=3, value=formula); c.font = Font(name=F, bold=bold, size=10, color=font.color); c.number_format = fmt
    w = T.cell(row=tr, column=4, value=words); w.font = NOTE; w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 5): T.cell(row=tr, column=col).fill = fill
    tr += 1; return f"C{tr-1}"
def tsec(name):
    global tr
    T.cell(row=tr, column=2, value=name).font = BOLD
    for col in range(1, 5): T.cell(row=tr, column=col).fill = SECFILL
    tr += 1

tsec("Path A: from NVIDIA's disclosed sovereign AI revenue")
a1 = trow("A1", "NVIDIA sovereign AI revenue, FY2026", "=" + ASSUMP["nvda_sov_fy26"], ">$30B, tripled year on year.", USD1, GREEN)
a2 = trow("A2", "Growth to calendar 2026", "=" + ASSUMP["nvda_sov_growth"], "Assumption 40%.", PCT, GREEN)
a3 = trow("A3", "NVIDIA sovereign AI revenue, 2026", f"={a1}*(1+{a2})", "A1 x (1 + A2).", USD1)
a4 = trow("A4", "NVIDIA share of sovereign accelerator spend", "=" + ASSUMP["nvda_share"], "Assumption 85%.", PCT, GREEN)
a5 = trow("A5", "All-vendor accelerator and system spend by sovereigns, 2026", f"={a3}/{a4}", "A3 / A4.", USD1)
a6 = trow("A6", "Server share of data center capex", "=" + ASSUMP["server_share"], "Carnegie: 64%.", PCT, GREEN)
a7 = trow("A7", "Sovereign AI data center capex, all-in, 2026 (Path A)", f"={a5}/{a6}", "A5 / A6. Grosses GPU systems up to the whole facility (network, storage, building, power).", USD1, BLACK, True, KEYFILL)
a8 = trow("A8", "Implied gigawatts energised at Epoch's $38B per GW", f"={a7}/{ASSUMP['capex_gw']}", "A7 / $38B. Note NVIDIA books revenue when systems ship, often a year before power is live.", '0.00')

tsec("Path B: from analyst market estimates")
b1 = trow("B1", "Sovereign AI market, 2026 (MarketsandMarkets, grown from 2025)", f"={ASSUMP['mm_2025']}*(1+{ASSUMP['mm_cagr']})", "$40.0B x 1.206.", USD1)
b2 = trow("B2", "Sovereign AI market, 2026 (Grand View Research)", "=" + ASSUMP["gvr_2026"], "$48.9B.", USD1, GREEN)
b3 = trow("B3", "Hardware share", "=" + ASSUMP["gvr_hw"], "48%.", PCT, GREEN)
b4 = trow("B4", "Sovereign AI hardware, 2026, average of B1 and B2 times B3", f"=AVERAGE({b1},{b2})*{b3}", "AVERAGE(B1, B2) x B3.", USD1)
b5 = trow("B5", "Sovereign AI infrastructure, 2026 (Precedence Research)", "=" + ASSUMP["prec_2026"], "$19.2B. Narrower definition: national data centers and government AI cloud.", USD1, GREEN)
b6 = trow("B6", "Sovereign AI infrastructure, 2026 (SNS Insider, grown at its own CAGR)", f"={ASSUMP['sns_2025']}*(91.99/17.81)^(1/10)", "$17.81B x ten-year CAGR implied by $91.99B in 2035.", USD1)
b7 = trow("B7", "Path B: average of the three infrastructure figures (B4, B5, B6)", f"=AVERAGE({b4},{b5},{b6})", "AVERAGE(B4, B5, B6). Analyst definitions differ; hardware-only in B4, infrastructure incl. cloud in B5 and B6.", USD1, BLACK, True, KEYFILL)

tsec("Path C: from public budgets with private leverage")
c1 = trow("C1", "EU gigafactories and AI Factories, annualised ($B)", f"=({ASSUMP['eu_giga']}+{ASSUMP['eu_factories']})/{ASSUMP['eu_years']}*{ASSUMP['eur_usd']}", "(EUR 10B + EUR 10B) / 5 years x 1.08.", USD2)
c2 = trow("C2", "Japan compute budget plus Noetra, annualised ($B)", f"={ASSUMP['jp_budget']}*{ASSUMP['jp_compute_share']}+{ASSUMP['jp_rubin_inv']}/5", "$7.9B x 40% + $6.7B / 5.", USD2)
c3 = trow("C3", "Korea government GPUs, annualised ($B)", f"={ASSUMP['kr_gpus_gov']}*{ASSUMP['gb300_chip']}*2/{ASSUMP['kr_years']}/1e9", "50,000 GPUs x ~$28.6k chip x 2 (system) / 3 years. Industrial fleets (SK, Hyundai, NAVER) are private and enter through leverage.", USD2)
c4 = trow("C4", "India, UK, France, Malaysia, Saudi NIF, Canada, annualised ($B)", f"={ASSUMP['in_budget']}/5+{ASSUMP['uk_budget']}*{ASSUMP['gbp_usd']}/{ASSUMP['uk_years']}+{ASSUMP['fr_bpi']}*{ASSUMP['eur_usd']}/{ASSUMP['fr_years']}+{ASSUMP['my_budget']}+{ASSUMP['sa_nif']}+{ASSUMP['ca_budget']}*{ASSUMP['cad_usd']}/{ASSUMP['ca_years']}", "Each programme's budget divided by its years, converted to USD.", USD2)
c5 = trow("C5", "Public AI compute budgets, 2026", f"={c1}+{c2}+{c3}+{c4}", "C1 + C2 + C3 + C4. Excludes Gulf sovereign wealth, which is equity rather than budget.", USD1)
c6 = trow("C6", "Private and sovereign-wealth capital per public dollar", "=" + ASSUMP["leverage"], "Assumption 2x (Commission targets 10x).", MULT, GREEN)
c7 = trow("C7", "Sovereign AI compute spend, 2026 (Path C)", f"={c5}*(1+{c6})", "C5 x (1 + C6).", USD1, BLACK, True, KEYFILL)

tsec("Reconciliation")
r1 = trow("R1", "Path A: NVIDIA sovereign revenue grossed up to the data center", f"={a7}", "", USD1)
r2 = trow("R2", "Path B: analyst infrastructure estimates", f"={b7}", "", USD1)
r3 = trow("R3", "Path C: public budgets with leverage", f"={c7}", "", USD1)
r4 = trow("R4", "Average of the three paths", f"=AVERAGE({r1},{r2},{r3})", "AVERAGE(R1:R3).", USD1, BLACK, True, KEYFILL)
r5 = trow("R5", "Spread (max minus min) over average", f"=(MAX({r1},{r2},{r3})-MIN({r1},{r2},{r3}))/{r4}", "", PCT)
r6 = trow("R6", "Bottom-up: capex energised in 2026, all-in", f"={BU['capex26']}", "From BottomUp summary.", USD1, GREEN, True)
r7 = trow("R7", "Bottom-up as share of top-down average", f"={r6}/{r4}", "R6 / R4.", PCT, BLACK, True, KEYFILL)
r8 = trow("R8", "Bottom-up gigawatts energised 2026", f"={BU['gw26']}", "", '0.00', GREEN)
r9 = trow("R9", "Path A implied gigawatts (A8) over bottom-up gigawatts", f"={a8}/{r8}", "A8 / R8. Above 1.0 means NVIDIA is shipping ahead of the power being ready, i.e. GPUs sit in warehouses or on racks waiting for grid connection.", MULT, BLACK, True, KEYFILL)
r10 = trow("R10", "Sovereign share of AI providers' $1T infrastructure spend", f"={r4}/{ASSUMP['providers_2026']}", "R4 / $1,000B.", PCT)
r11 = trow("R11", "Sovereign share of NVIDIA Data Center run-rate (A3 / annualised Q2 FY27 $89B)", f"={a3}/(89*4)", "A3 / $356B.", PCT)
T.freeze_panes = "A5"

# ---------------------------------------------------------------- GigawattTCO
G = wb.create_sheet("GigawattTCO")
title(G, "What a gigawatt costs to own, by country", "Epoch's US hyperscaler model, then adjusted for power price, PUE, delay and insurance. Owner's effective $ per GPU-hour at the bottom. $B unless noted.")
header(G, 4, [1, 2, 3, 4, 5, 6, 7, 8], ["Line", "Item", "United States", "Saudi Arabia", "South Korea", "Germany", "India", "Formula (in words)"])
widths(G, [6, 56, 14, 14, 14, 14, 14, 90])
gr = 5
def grow(n, label, fs, words, fmt=USD2, bold=False, fill=None, fonts=None):
    global gr
    G.cell(row=gr, column=1, value=n).font = NOTE
    G.cell(row=gr, column=2, value=label).font = BOLD if bold else BLACK
    for i, f in enumerate(fs):
        c = G.cell(row=gr, column=3 + i, value=f)
        fnt = fonts[i] if fonts else BLACK
        c.font = Font(name=F, bold=bold, size=10, color=fnt.color); c.number_format = fmt
    w = G.cell(row=gr, column=8, value=words); w.font = NOTE; w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 9): G.cell(row=gr, column=col).fill = fill
    gr += 1; return gr - 1
def gsec(name):
    global gr
    G.cell(row=gr, column=2, value=name).font = BOLD
    for col in range(1, 9): G.cell(row=gr, column=col).fill = SECFILL
    gr += 1
gsec("Inputs by country")
rP = grow("G1", "Industrial power price ($/kWh)", ["=" + ASSUMP["us_power"], "=" + ASSUMP["power_sa"], "=" + ASSUMP["power_kr"], "=" + ASSUMP["power_de"], "=" + ASSUMP["power_in"]], "From Assumptions.", '$0.000', fonts=[GREEN]*5)
rPUE = grow("G2", "PUE", [1.2, "=" + ASSUMP["pue_gulf"], "=" + ASSUMP["pue_eu"], "=" + ASSUMP["pue_eu"], 1.3], "US and India my figures; others from Assumptions.", '0.00', fonts=[BLUE, GREEN, GREEN, GREEN, BLUE])
rD = grow("G3", "Extra years to energise vs US", [0, 0, 0.5, "=" + ASSUMP["delay_de"], 1.0], "Carnegie timelines: Gulf fastest, Germany slowest. Korea and India my figures.", '0.0', fonts=[BLUE, BLUE, BLUE, GREEN, BLUE])
rI = grow("G4", "Insurance and country-risk premium (share of capex per year)", [0.005, "=" + ASSUMP["war_ins"], 0.005, 0.005, 0.01], "Gulf war-risk insurance 5% after April 2026 (Semafor); others nominal.", PCT, fonts=[BLUE, GREEN, BLUE, BLUE, BLUE])
rCF = grow("G5", "Carnegie lifecycle value index vs US (reference only)", ["=" + ASSUMP["cf_us"], "=" + ASSUMP["cf_uae"], "=" + ASSUMP["cf_kr"], "=" + ASSUMP["cf_de"], "=" + ASSUMP["cf_in"]], "Saudi uses the UAE figure. Independent check on the rows below.", PCT, fonts=[GREEN]*5)
gsec("Cost of 1 GW IT load, $B")
rCap = grow("G6", "Up-front capex (Epoch, same hardware everywhere)", ["=" + ASSUMP["capex_gw"]] * 5, "$38B. GPUs are priced the same worldwide (Carnegie), so capex barely moves by country.", USD1, fonts=[GREEN]*5)
rEn = grow("G7", "Energy per year", [f"=1000000*{get_column_letter(3+i)}{rPUE}*{ASSUMP['hours']}*{ASSUMP['util']}*{get_column_letter(3+i)}{rP}/1e9" for i in range(5)], "1,000,000 kW x PUE x 8,760 h x utilisation x $/kWh / 1e9. Epoch's US figure is $0.59B at full load.", USD2)
rOp = grow("G8", "Other operating cost per year (Epoch opex less energy)", [f"={ASSUMP['opex_gw']}-{ASSUMP['energy_gw_yr']}"] * 5, "$0.9B - $0.59B = $0.31B (labour, maintenance, taxes, water).", USD2)
rIns = grow("G9", "Insurance and risk per year", [f"={get_column_letter(3+i)}{rCap}*{get_column_letter(3+i)}{rI}" for i in range(5)], "G6 x G4.", USD2)
rAnn = grow("G10", "Annualised capex (servers over IT life, rest over 15 years)", [f"={get_column_letter(3+i)}{rCap}*{ASSUMP['server_share']}/{ASSUMP['server_life']}+{get_column_letter(3+i)}{rCap}*(1-{ASSUMP['server_share']})/{ASSUMP['facility_life']}" for i in range(5)], "Capex x 64% / 4 years + capex x 36% / 10 years. Epoch's own annualisation gives $7.6B of capex-related cost per year.", USD2)
rTCO = grow("G11", "Annual total cost of ownership", [f"={get_column_letter(3+i)}{rAnn}+{get_column_letter(3+i)}{rEn}+{get_column_letter(3+i)}{rOp}+{get_column_letter(3+i)}{rIns}" for i in range(5)], "G10 + G7 + G8 + G9. Epoch's US figure is $8.5B.", USD2, True, KEYFILL)
rDel = grow("G12", "Lifecycle value lost to delay (share)", [f"={get_column_letter(3+i)}{rD}*{ASSUMP['delay_1y']}" for i in range(5)], "Extra years x 5.5% per year (Carnegie, linearised). A 1.5-year German delay costs more than doubling its power price.", PCT)
rDelB = grow("G13", "Delay cost, $B over a 10-year life (share x 10 x TCO)", [f"={get_column_letter(3+i)}{rDel}*10*{get_column_letter(3+i)}{rTCO}" for i in range(5)], "G12 x 10 years x G11.", USD2)
rEnS = grow("G14", "Energy as share of annual TCO", [f"={get_column_letter(3+i)}{rEn}/{get_column_letter(3+i)}{rTCO}" for i in range(5)], "G7 / G11. Even in Germany, energy is a minority of cost. Time to power matters more than the price of power.", PCT)
gsec("Owner's unit economics")
rGPU = grow("G15", "GPUs in 1 GW at NVL72 density", [f"=1000000/{ASSUMP['rack_kw']}*{ASSUMP['rack_gpus']}"] * 5, "1,000,000 kW / 120 kW x 72.", NUM)
rHrs = grow("G16", "Used GPU-hours per year", [f"={get_column_letter(3+i)}{rGPU}*{ASSUMP['hours']}*{ASSUMP['util']}" for i in range(5)], "G15 x 8,760 x 70%.", NUM)
rPh = grow("G17", "Owner's effective cost per used GPU-hour", [f"={get_column_letter(3+i)}{rTCO}*1e9/{get_column_letter(3+i)}{rHrs}" for i in range(5)], "G11 x 1e9 / G16.", USD2, True, KEYFILL)
rPh2 = grow("G18", "Same, including delay cost spread over 10 years", [f"=({get_column_letter(3+i)}{rTCO}+{get_column_letter(3+i)}{rDelB}/10)*1e9/{get_column_letter(3+i)}{rHrs}" for i in range(5)], "(G11 + G13 / 10) x 1e9 / G16.", USD2, True)
rCapCh = grow("G18b", "Cost of capital on capex per year (12% hurdle), $B", [f"={get_column_letter(3+i)}{rCap}*{ASSUMP['hurdle']}" for i in range(5)], "G6 x 12%. What the owner must earn above TCO to justify the capital.", USD2)
rPh3 = grow("G18c", "Cost per used GPU-hour incl. 12% cost of capital", [f"=({get_column_letter(3+i)}{rTCO}+{get_column_letter(3+i)}{rCapCh})*1e9/{get_column_letter(3+i)}{rHrs}" for i in range(5)], "(G11 + G18b) x 1e9 / G16. The break-even rental price for a financed owner.", USD2, True, KEYFILL)
rVs = grow("G19", "Owner's cost as share of GB200 on-demand rental ($10.50)", [f"={get_column_letter(3+i)}{rPh}/{ASSUMP['gb200_hr']}" for i in range(5)], "G17 / $10.50 (CoreWeave, May 2026).", PCT)
rVs2 = grow("G20", "Owner's cost as share of Nebius B200 ($8.50)", [f"={get_column_letter(3+i)}{rPh}/{ASSUMP['p_nebius_b200']}" for i in range(5)], "G17 / $8.50.", PCT)
rCapGPU = grow("G21", "Capex per GPU, all-in ($)", [f"={get_column_letter(3+i)}{rCap}*1e9/{get_column_letter(3+i)}{rGPU}" for i in range(5)], "G6 x 1e9 / G15. Compare the GB300 export-approval value of ~$28.6k per chip and street B300 ~$53k.", USD)
G.freeze_panes = "C5"
GT = {"us_ph3": f"GigawattTCO!$C${rPh3}", "sa_ph3": f"GigawattTCO!$D${rPh3}", "kr_ph3": f"GigawattTCO!$E${rPh3}", "de_ph3": f"GigawattTCO!$F${rPh3}", "in_ph3": f"GigawattTCO!$G${rPh3}", "us_ph": f"GigawattTCO!$C${rPh}", "sa_ph": f"GigawattTCO!$D${rPh}", "kr_ph": f"GigawattTCO!$E${rPh}", "de_ph": f"GigawattTCO!$F${rPh}", "in_ph": f"GigawattTCO!$G${rPh}", "de_ph2": f"GigawattTCO!$F${rPh2}", "us_tco": f"GigawattTCO!$C${rTCO}", "gpu_gw": f"GigawattTCO!$C${rGPU}", "de_del": f"GigawattTCO!$F${rDelB}", "de_en": f"GigawattTCO!$F${rEn}", "us_en": f"GigawattTCO!$C${rEn}", "cap_gpu": f"GigawattTCO!$C${rCapGPU}", "sa_tco": f"GigawattTCO!$D${rTCO}", "de_tco": f"GigawattTCO!$F${rTCO}"}

# ---------------------------------------------------------------- Configs
K = wb.create_sheet("Configs")
title(K, "Four reference configurations, one per quadrant", "Specified and costed on the GigawattTCO rates. Columns are comparable.")
header(K, 4, [1, 2, 3, 4, 5, 6], ["Line", "Item", "Compute exporter campus", "Industrial AI factory", "EU gigafactory", "Subsidised national GPU pool"])
widths(K, [6, 58, 20, 20, 20, 20])
K.cell(row=5, column=2, value="Who builds it").font = BOLD
for col, t in zip((3, 4, 5, 6), ["HUMAIN with xAI, G42 with OpenAI and Oracle", "Hyundai, SK, SoftBank's Noetra, Deutsche Telekom", "EuroHPC consortium under InvestAI", "IndiaAI, Malaysia sovereign cloud, Thailand"]):
    c = K.cell(row=5, column=col, value=t); c.font = NOTE; c.alignment = Alignment(wrap_text=True, vertical="top")
K.row_dimensions[5].height = 40
kr = 6
def krow(n, label, fs, fmt=NUM, bold=False, fill=None, fonts=None):
    global kr
    K.cell(row=kr, column=1, value=n).font = NOTE
    K.cell(row=kr, column=2, value=label).font = BOLD if bold else BLACK
    for i, f in enumerate(fs):
        c = K.cell(row=kr, column=3 + i, value=f); fnt = fonts[i] if fonts else BLACK
        c.font = Font(name=F, bold=bold, size=10, color=fnt.color); c.number_format = fmt; c.alignment = Alignment(horizontal="right", wrap_text=True)
    if fill:
        for col in range(1, 7): K.cell(row=kr, column=col).fill = fill
    kr += 1; return kr - 1
def ksec(name):
    global kr
    K.cell(row=kr, column=2, value=name).font = BOLD
    for col in range(1, 7): K.cell(row=kr, column=col).fill = SECFILL
    kr += 1
ksec("Specification")
kMW = krow("K1", "IT load, MW", [500, 100, 190, 20], NUM, fonts=[BLUE]*4)
kRack = krow("K2", "Rack type (kW)", ["=" + ASSUMP["gb300_rack_kw"], "=" + ASSUMP["rack_kw"], "=" + ASSUMP["rack_kw"], 14.3], NUM1, fonts=[GREEN, GREEN, GREEN, BLUE])
kGpr = krow("K3", "GPUs per rack or node", [72, 72, 72, 8], NUM, fonts=[BLUE]*4)
kG = krow("K4", "GPUs", [f"={get_column_letter(3+i)}{kMW}*1000/{get_column_letter(3+i)}{kRack}*{get_column_letter(3+i)}{kGpr}" for i in range(4)], NUM, True)
kCool = krow("K5", "Cooling", ["Liquid, closed loop", "Liquid", "Liquid", "Air"], '@', fonts=[BLUE]*4)
kPow = krow("K6", "Power source", ["Grid plus dedicated gas and solar PPAs", "Grid, 13c industrial", "Grid, 18c, 1.5-year interconnect queue", "Grid, 9c"], '@', fonts=[BLUE]*4)
kWho = krow("K7", "Who pays", ["Sovereign wealth fund equity plus project debt", "Chaebol or telco balance sheet with state subsidy", "EU and national budgets plus EIB, 10x private leverage target", "State budget subsidises the hour; hyperscaler or neocloud builds"], '@', fonts=[BLUE]*4)
kCust = krow("K8", "Anchor customer", ["xAI, OpenAI, foreign AI labs renting capacity", "Own group companies: automotive, semis, robotics", "Mistral, Black Forest Labs, Helsing, research", "Startups and universities at a subsidised rate"], '@', fonts=[BLUE]*4)
ksec("Cost, $B")
kPr = krow("K9", "Power price ($/kWh)", ["=" + ASSUMP["power_sa"], "=" + ASSUMP["power_kr"], "=" + ASSUMP["power_de"], "=" + ASSUMP["power_in"]], '$0.000', fonts=[GREEN]*4)
kPue = krow("K10", "PUE", ["=" + ASSUMP["pue_gulf"], "=" + ASSUMP["pue_eu"], "=" + ASSUMP["pue_eu"], 1.3], '0.00', fonts=[GREEN, GREEN, GREEN, BLUE])
kCap = krow("K11", "Up-front capex ($B)", [f"={get_column_letter(3+i)}{kMW}/1000*{ASSUMP['capex_gw']}" for i in range(4)], USD2, True)
kEn = krow("K12", "Energy per year ($B)", [f"={get_column_letter(3+i)}{kMW}*1000*{get_column_letter(3+i)}{kPue}*{ASSUMP['hours']}*{ASSUMP['util']}*{get_column_letter(3+i)}{kPr}/1e9" for i in range(4)], USD2)
kOp = krow("K13", "Other opex per year ($B)", [f"={get_column_letter(3+i)}{kMW}/1000*({ASSUMP['opex_gw']}-{ASSUMP['energy_gw_yr']})" for i in range(4)], USD2)
kIns = krow("K14", "Insurance and risk per year ($B)", [f"={get_column_letter(3+i)}{kCap}*{v}" for i, v in enumerate([ASSUMP['war_ins'], 0.005, 0.005, 0.01])], USD2)
kAnn = krow("K15", "Annualised capex ($B)", [f"={get_column_letter(3+i)}{kCap}*{ASSUMP['server_share']}/{ASSUMP['server_life']}+{get_column_letter(3+i)}{kCap}*(1-{ASSUMP['server_share']})/{ASSUMP['facility_life']}" for i in range(4)], USD2)
kTCO = krow("K16", "Annual TCO ($B)", [f"={get_column_letter(3+i)}{kAnn}+{get_column_letter(3+i)}{kEn}+{get_column_letter(3+i)}{kOp}+{get_column_letter(3+i)}{kIns}" for i in range(4)], USD2, True, KEYFILL)
kDel = krow("K17", "Years of delay vs US", [0, 0.5, "=" + ASSUMP["delay_de"], 1.0], '0.0', fonts=[BLUE, BLUE, GREEN, BLUE])
kDelB = krow("K18", "Delay cost over life ($B)", [f"={get_column_letter(3+i)}{kDel}*{ASSUMP['delay_1y']}*10*{get_column_letter(3+i)}{kTCO}" for i in range(4)], USD2)
ksec("Unit economics")
kUt = krow("K19", "Utilisation", [0.8, "=" + ASSUMP["util"], 0.6, 0.5], PCT, fonts=[BLUE, GREEN, BLUE, BLUE])
kHr = krow("K20", "Used GPU-hours per year", [f"={get_column_letter(3+i)}{kG}*{ASSUMP['hours']}*{get_column_letter(3+i)}{kUt}" for i in range(4)], NUM)
kPh = krow("K21", "Owner's cost per used GPU-hour ($)", [f"={get_column_letter(3+i)}{kTCO}*1e9/{get_column_letter(3+i)}{kHr}" for i in range(4)], USD2, True, KEYFILL)
kPh3 = krow("K21b", "Break-even price incl. 12% cost of capital ($ per GPU-hour)", [f"=({get_column_letter(3+i)}{kTCO}+{get_column_letter(3+i)}{kCap}*{ASSUMP['hurdle']})*1e9/{get_column_letter(3+i)}{kHr}" for i in range(4)], USD2, True)
kSell = krow("K22", "Price charged per GPU-hour ($)", [f"={ASSUMP['gb200_hr']}*0.6", "n/a (internal)", "=" + ASSUMP["p_scaleway"], "=" + ASSUMP["p_india_sub"]], USD2, fonts=[BLUE, BLUE, GREEN, GREEN])
kMar = krow("K23", "Gross margin or subsidy per GPU-hour ($)", [f"=C{kSell}-C{kPh}", "n/a", f"=E{kSell}-E{kPh}", f"=F{kSell}-F{kPh}"], USD2, True)
kSub = krow("K24", "Annual subsidy or margin at the hours above ($M)", [f"=C{kMar}*C{kHr}/1e6", "n/a", f"=E{kMar}*E{kHr}/1e6", f"=F{kMar}*F{kHr}/1e6"], USD1, True, KEYFILL)
kFit = krow("K25", "Reading", ["Sells capacity abroad at a margin while power is cheap and chips are allocated. Exposed to export policy and war risk.", "Cost centre for the group. Justified by research velocity and supply-chain sovereignty, not by $/hour.", "Priced to European sovereign-cloud rates it cannot reach at 60% utilisation and German power. Needs the 10x leverage and an anchor tenant.", "Every subsidised hour costs the state the gap. Cheapest way to give startups compute, most expensive per GPU."], '@', fonts=[BLUE]*4)
K.cell(row=kr, column=2, value="Formulas: GPUs = MW x 1,000 / rack kW x GPUs per rack. Capex = MW / 1,000 x $38B. Energy = MW x 1,000 x PUE x 8,760 x utilisation x $/kWh. TCO = annualised capex + energy + other opex + insurance. $/GPU-hour = TCO / (GPUs x 8,760 x utilisation). Margin or subsidy = price - cost.").font = NOTE
K.freeze_panes = "C6"
KC = {}
for code, r_ in [("K21b", kPh3), ("K4", kG), ("K11", kCap), ("K16", kTCO), ("K18", kDelB), ("K19", kUt), ("K21", kPh), ("K22", kSell), ("K23", kMar), ("K24", kSub)]:
    KC[code] = r_

# ---------------------------------------------------------------- Pricing
P = wb.create_sheet("Pricing")
title(P, "Pricing: GPU-hour by route and region, subsidy gaps, and where the margin sits", "List prices read Sep-Oct 2026. Owner costs from GigawattTCO.")
header(P, 4, [1, 2, 3, 4], ["Line", "Item", "$ per GPU-hour", "Note"])
widths(P, [6, 70, 16, 100])
pr = 5
def prow(n, label, f, note, fmt=USD2, font=BLACK, bold=False, fill=None):
    global pr
    P.cell(row=pr, column=1, value=n).font = NOTE
    P.cell(row=pr, column=2, value=label).font = BOLD if bold else BLACK
    c = P.cell(row=pr, column=3, value=f); c.font = Font(name=F, bold=bold, size=10, color=font.color); c.number_format = fmt
    w = P.cell(row=pr, column=4, value=note); w.font = NOTE; w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 5): P.cell(row=pr, column=col).fill = fill
    pr += 1; return f"C{pr-1}"
def psec(name):
    global pr
    P.cell(row=pr, column=2, value=name).font = BOLD
    for col in range(1, 5): P.cell(row=pr, column=col).fill = SECFILL
    pr += 1
psec("Owner's cost (from GigawattTCO, GB200 class, 70% utilisation)")
o1 = prow("P1", "Own a gigawatt in Saudi Arabia", "=" + GT["sa_ph"], "", USD2, GREEN, True)
o2 = prow("P2", "Own a gigawatt in the United States", "=" + GT["us_ph"], "", USD2, GREEN)
o3 = prow("P3", "Own a gigawatt in South Korea", "=" + GT["kr_ph"], "", USD2, GREEN)
o4 = prow("P4", "Own a gigawatt in Germany", "=" + GT["de_ph"], "", USD2, GREEN)
o5 = prow("P5", "Own a gigawatt in Germany incl. 1.5-year delay cost", "=" + GT["de_ph2"], "", USD2, GREEN)
o6 = prow("P5b", "Break-even rental price, US owner incl. 12% cost of capital", "=" + GT["us_ph3"], "From GigawattTCO G18c.", USD2, GREEN, True)
o7 = prow("P5c", "Break-even rental price, Saudi owner incl. 12% cost of capital", "=" + GT["sa_ph3"], "", USD2, GREEN)
o8 = prow("P5d", "Break-even rental price, German owner incl. 12% cost of capital", "=" + GT["de_ph3"], "", USD2, GREEN)
psec("Rent, Blackwell class")
r_1 = prow("P6", "CoreWeave GB200 on-demand (lowest, May 2026)", "=" + ASSUMP["gb200_hr"], "Carnegie / getdeploying.", USD2, GREEN)
r_2 = prow("P7", "Nebius B200", "=" + ASSUMP["p_nebius_b200"], "14 Oct 2026.", USD2, GREEN)
r_3 = prow("P8", "Verda B200, Finland", "=" + ASSUMP["p_verda_b200"], "Sep 2026.", USD2, GREEN)
psec("Rent, Hopper class (what most sovereign clouds actually sell)")
h1 = prow("P9", "Azure H100, West Europe", "=" + ASSUMP["p_azure_eu"], "Highest list in Europe.", USD2, GREEN)
h2 = prow("P10", "AWS Japan H100", "=" + ASSUMP["p_aws_jp"], "", USD2, GREEN)
h3 = prow("P11", "AWS H100, Stockholm", "=" + ASSUMP["p_aws_eu"], "", USD2, GREEN)
h4 = prow("P12", "CoreWeave HGX H100, Europe", "=" + ASSUMP["p_coreweave_h100"], "", USD2, GREEN)
h5 = prow("P13", "Sakura Internet H100, Japan (METI-subsidised operator)", "=" + ASSUMP["p_sakura"], "JPY 990.", USD2, GREEN)
h6 = prow("P14", "Nebius H100 SXM", "=" + ASSUMP["p_nebius_h100"], "", USD2, GREEN)
h7 = prow("P15", "Scaleway H100 PCIe, Paris", "=" + ASSUMP["p_scaleway"], "EUR 2.87.", USD2, GREEN)
h8 = prow("P16", "OVHcloud H100, Gravelines", "=" + ASSUMP["p_ovh"], "Won EUR 180M Commission sovereign-cloud contract.", USD2, GREEN)
h9 = prow("P17", "Global H100 on-demand median (NVIDIA index, Jun 2026)", "=" + ASSUMP["p_h100_median"], "", USD2, GREEN)
h10 = prow("P18", "IndiaAI subsidised rate", "=" + ASSUMP["p_india_sub"], "INR 150.", USD2, GREEN, True)
psec("Spreads")
s1 = prow("P19", "European sovereign cloud discount to Azure West Europe (P16 / P9)", f"={h8}/{h1}", "OVHcloud as share of Azure list. European sovereign clouds compete on price, not just jurisdiction.", PCT, BLACK, True, KEYFILL)
s2 = prow("P20", "Hyperscaler Japan premium over METI-subsidised Sakura (P10 / P13)", f"={h2}/{h5}", "", MULT)
s3 = prow("P21", "India subsidy gap per hour vs commercial rate (P18 vs INR 300)", f"={ASSUMP['p_india_com']}-{h10}", "Commercial $3.40 less subsidised $1.70.", USD2)
s4 = prow("P22", "India subsidy gap vs owner's cost of a US gigawatt (P2 - P18)", f"={o2}-{h10}", "What the state eats per hour if the pool costs what a US hyperscaler pays.", USD2, BLACK, True, KEYFILL)
s5 = prow("P23", "Annual subsidy on 45,000 GPUs at 50% utilisation ($M)", f"={s4}*45000*{ASSUMP['hours']}*0.5/1e6", "P22 x 45,000 x 8,760 x 50% / 1e6. Against IndiaAI's ~$1.2B five-year budget.", USD1, BLACK, True, KEYFILL)
s6 = prow("P24", "Saudi owner's cost as share of CoreWeave GB200 rental (P1 / P6)", f"={o1}/{r_1}", "The export margin available to a compute exporter before financing cost.", PCT, BLACK, True, KEYFILL)
s7b_placeholder = None
s7 = prow("P25", "Germany owner's cost incl. delay as share of Scaleway's price (P5 / P15)", f"={o5}/{h7}", "Below 100% means a German gigafactory can undercut existing European sovereign clouds on cash cost. Add the 12% cost of capital (P5d) and the comparison flips.", PCT, BLACK, True, KEYFILL)
s8 = prow("P25b", "German break-even incl. cost of capital as share of Scaleway's price (P5d / P15)", f"={o8}/{h7}", "Above 100% means the gigafactory cannot earn its cost of capital selling at today's European sovereign-cloud prices.", PCT, BLACK, True, KEYFILL)
s9 = prow("P25c", "Saudi break-even incl. cost of capital as share of CoreWeave GB200 (P5c / P6)", f"={o7}/{r_1}", "The exporter still clears its capital charge with room to spare.", PCT, BLACK, True, KEYFILL)
psec("Where the margin sits")
m1 = prow("P26", "GB300 chip value in US export approvals ($ per chip)", "=" + ASSUMP["gb300_chip"], "35,000 chips ~ $1B.", USD, GREEN)
m2 = prow("P27", "NVIDIA cost of goods per chip at 75% gross margin ($)", f"={m1}*(1-{ASSUMP['nvda_gm']})", "P26 x 25%.", USD)
m3 = prow("P28", "All-in capex per GPU in a gigawatt ($)", "=" + GT["cap_gpu"], "From GigawattTCO: $38B / GPUs in 1 GW.", USD, GREEN)
m4 = prow("P29", "Chip as share of all-in capex per GPU (P26 / P28)", f"={m1}/{m3}", "The rest is the server, rack, network, storage, building and power plant.", PCT)
m5 = prow("P30", "Reading", "", "In a sovereign deal the chip is under half the money. The state controls the land, power and permits, which is why the Gulf can clear a gigawatt in 2.5 years and Northumberland takes nine. Pricing power in this market belongs to whoever controls time to power, and today that is the Gulf and NVIDIA's allocation desk, not the cloud.", '@')
P.freeze_panes = "A5"
PR = {"s8": s8, "s9": s9, "o6": o6, "o7": o7, "o8": o8, "s1": s1, "s4": s4, "s5": s5, "s6": s6, "s7": s7, "m4": m4, "o1": o1, "o2": o2, "o4": o4, "o5": o5}

# ---------------------------------------------------------------- Value
V = wb.create_sheet("Value")
title(V, "Value: what sovereign compute is worth", "Macro claims are large and soft. The usable tests are ratios: value per dollar invested, revenue per compute dollar, and whether the owner can cover cost of capital.")
header(V, 4, [1, 2, 3, 4], ["Line", "Item", "Value", "Formula (in words)"])
widths(V, [6, 72, 18, 100])
vr = 5
def vrow(n, label, f, words, fmt=USD1, font=BLACK, bold=False, fill=None):
    global vr
    V.cell(row=vr, column=1, value=n).font = NOTE
    V.cell(row=vr, column=2, value=label).font = BOLD if bold else BLACK
    c = V.cell(row=vr, column=3, value=f); c.font = Font(name=F, bold=bold, size=10, color=font.color); c.number_format = fmt
    w = V.cell(row=vr, column=4, value=words); w.font = NOTE; w.alignment = Alignment(wrap_text=True, vertical="top")
    if fill:
        for col in range(1, 5): V.cell(row=vr, column=col).fill = fill
    vr += 1; return f"C{vr-1}"
def vsec(name):
    global vr
    V.cell(row=vr, column=2, value=name).font = BOLD
    for col in range(1, 5): V.cell(row=vr, column=col).fill = SECFILL
    vr += 1
vsec("A. Macro: GDP uplift that depends on in-country compute")
va1 = vrow("VA1", "PwC: AI uplift to global GDP by 2030 (share)", "=" + ASSUMP["pwc_pct"], "14%.", PCT, GREEN)
va2 = vrow("VA2", "Share of that uplift requiring in-country compute", "=" + ASSUMP["ai_share_gdp"], "Assumption 25%.", PCT, GREEN)
va3 = vrow("VA3", "Saudi Arabia: annual GDP uplift at stake ($B)", f"={ASSUMP['gdp_sa']}*1000*{va1}*{va2}", "GDP x 14% x 25%.", USD1)
va4 = vrow("VA4", "Saudi Arabia: HUMAIN announced capex to 2030 ($B)", f"=1.9*{ASSUMP['capex_gw']}", "1.9 GW x $38B.", USD1)
va5 = vrow("VA5", "Years of uplift to pay back the capex", f"={va4}/{va3}", "VA4 / VA3.", '0.0', BLACK, True, KEYFILL)
va6 = vrow("VA6", "EU: annual GDP uplift at stake ($B)", f"={ASSUMP['gdp_eu']}*1000*{va1}*{va2}", "", USD1)
va7 = vrow("VA7", "EU: gigafactories plus AI factories capex ($B)", f"=1.7*{ASSUMP['capex_gw']}", "1.7 GW announced x $38B.", USD1)
va8 = vrow("VA8", "Years of uplift to pay back", f"={va7}/{va6}", "", '0.00', BLACK, True, KEYFILL)
va9 = vrow("VA9", "India: annual GDP uplift at stake ($B)", f"={ASSUMP['gdp_in']}*1000*{va1}*{va2}", "", USD1)
va10 = vrow("VA10", "India: IndiaAI capex equivalent ($B)", f"=0.09*{ASSUMP['capex_gw']}", "90 MW x $38B.", USD1)
va11 = vrow("VA11", "Years of uplift to pay back", f"={va10}/{va9}", "A rounding error. India's constraint is not capital but power, chips and talent.", '0.00', BLACK, True, KEYFILL)
va12 = vrow("VA12", "Thailand: annual GDP uplift at stake ($B)", f"={ASSUMP['gdp_th']}*1000*{va1}*{va2}", "", USD1)
va13 = vrow("VA13", "Thailand: Gulf Development 2 GW plan capex ($B)", f"=2*{ASSUMP['capex_gw']}", "2 GW x $38B. Far beyond domestic need, so it only works as an export or hyperscaler-tenant play.", USD1)
va14 = vrow("VA14", "Years of uplift to pay back", f"={va13}/{va12}", "", '0.0', BLACK, True, KEYFILL)
vsec("B. Micro: value per dollar of infrastructure")
vb1 = vrow("VB1", "Google Bangkok region: economic value claimed per dollar invested", f"={ASSUMP['th_value']}/{ASSUMP['th_invest']}", "$41B / $1B (Public First for Google). A vendor-commissioned multiplier.", MULT, BLACK, True)
vb2 = vrow("VB2", "Frontier lab revenue per dollar of compute (Anthropic)", "=" + ASSUMP["anthropic_ratio"], "$1.70. The best disclosed return on compute in the industry.", USD2, GREEN, True)
vb3 = vrow("VB3", "Gartner: revenue needed vs revenue available to AI data center owners through 2030", f"={ASSUMP['gartner_rev_have']}/{ASSUMP['gartner_rev_need']}", "$3.6T / $8.5T. Owners as a class cover 42% of what a 12% return requires.", PCT, BLACK, True, KEYFILL)
vb4 = vrow("VB4", "Infrastructure exposed to impairment by 2030 ($T)", "=" + ASSUMP["gartner_impair"], ">$3.5T.", USD1, GREEN)
vsec("C. Owner's economics for a sovereign gigawatt")
vc1 = vrow("VC1", "Saudi gigawatt annual TCO ($B)", "=" + GT["sa_tco"], "", USD2, GREEN)
vc2 = vrow("VC2", "Revenue if sold at 60% of CoreWeave GB200 rate, 80% utilisation ($B)", f"={ASSUMP['gb200_hr']}*0.6*{GT['gpu_gw']}*{ASSUMP['hours']}*0.8/1e9", "$10.50 x 60% x GPUs in 1 GW x 8,760 x 80% / 1e9.", USD2)
vc3 = vrow("VC3", "Operating margin before financing", f"=({vc2}-{vc1})/{vc2}", "(VC2 - VC1) / VC2.", PCT, BLACK, True, KEYFILL)
vc3b = vrow("VC3b", "Operating margin after 12% cost of capital on $38B", f"=({vc2}-{vc1}-{ASSUMP['capex_gw']}*{ASSUMP['hurdle']})/{vc2}", "(VC2 - VC1 - $38B x 12%) / VC2.", PCT, BLACK, True, KEYFILL)
vc4 = vrow("VC4", "German gigawatt annual TCO ($B)", "=" + GT["de_tco"], "", USD2, GREEN)
vc5 = vrow("VC5", "Revenue at Scaleway's Hopper-class price, 60% utilisation ($B)", f"={ASSUMP['p_scaleway']}*{GT['gpu_gw']}*{ASSUMP['hours']}*0.6/1e9", "EUR 2.87 x GPUs x 8,760 x 60% / 1e9. If the gigafactory has to compete with existing European clouds on price.", USD2)
vc6 = vrow("VC6", "Operating margin before financing", f"=({vc5}-{vc4})/{vc5}", "Negative means the gigafactory cannot be run as a cloud business at today's European prices. It is a research instrument and must be valued as one.", PCT, BLACK, True, KEYFILL)
vc6b = vrow("VC6b", "Operating margin after 12% cost of capital", f"=({vc5}-{vc4}-{ASSUMP['capex_gw']}*{ASSUMP['hurdle']})/{vc5}", "(VC5 - VC4 - $38B x 12%) / VC5. Negative means a European gigafactory cannot be run as a cloud business at today's European prices. It is a research instrument and must be valued as one.", PCT, BLACK, True, KEYFILL)
vc7 = vrow("VC7", "Revenue needed to cover German TCO plus 12% cost of capital at 60% utilisation, $ per GPU-hour", f"=({vc4}+{ASSUMP['capex_gw']}*{ASSUMP['hurdle']})*1e9/({GT['gpu_gw']}*{ASSUMP['hours']}*0.6)", "(VC4 + $38B x 12%) x 1e9 / used hours. The Blackwell-class price a gigafactory would have to charge.", USD2, BLACK, True)
vsec("D. The strategic value nobody prices")
vrow("VD1", "Reading", "", "Three things do not appear in the ratios above. Optionality: a state with its own gigawatt can run a national model when a foreign API is cut off. Industrial pull-through: Korea's 260,000 GPUs are for Hyundai and SK's products, not for selling compute. Bargaining power: HUMAIN's 35,000-chip export approval and $23B of agreements were negotiated at head-of-state level. These are real and they explain why the capital keeps coming after the spreadsheet says stop.", '@')
V.freeze_panes = "A5"
VV = {"vc3b": vc3b, "vc6b": vc6b, "va5": va5, "va8": va8, "va11": va11, "va14": va14, "vb1": vb1, "vb3": vb3, "vc3": vc3, "vc6": vc6, "vc7": vc7, "va3": va3, "va4": va4}

# ---------------------------------------------------------------- Sources
S = wb.create_sheet("Sources")
title(S, "Sources", "Numbering matches the article at esthercs.dev/research/sovereign-ai.")
header(S, 4, [1, 2, 3, 4], ["No.", "Source", "Date", "URL"])
widths(S, [6, 92, 16, 110])
SOURCES = [
    ("Fortune, Nvidia's European spending spree is about much more than Europe (sovereign revenue >$30B FY26, $8B into Europe, Nscale)", "15 Oct 2026", "https://fortune.com/2026/10/15/nvidias-european-spending-spree-is-about-much-more-than-europe/"),
    ("Air Street Capital, State of AI Report 2026 (NVIDIA >$30B sovereign AI revenue FY2026; hyperscaler capex ~$733B)", "Oct 2026", "https://www.stateof.ai/"),
    ("vision2030.ai, HUMAIN's AI Infrastructure Machine: 600,000 GPUs, $77 Billion, 6.6 GW", "2026", "https://vision2030.ai/analysis/humain-ai-infrastructure/"),
    ("Saudi Market Research Consulting, HUMAIN Saudi Arabia: A Bold Sovereign AI Push", "2025-2026", "https://saudimarketresearchconsulting.com/insights/articles/humain-saudi-arabia-a-bold-sovereign-ai-push-reshaping-infrastructure-and-markets"),
    ("Qualcomm, HUMAIN and Qualcomm to deploy AI infrastructure in Saudi Arabia (200 MW)", "Oct 2025", "https://www.qualcomm.com/news/releases/2025/10/humain-and-qualcomm-to-deploy-ai-infrastructure-in-saudi-arabia-"),
    ("US Department of Commerce, Statement on UAE and Saudi Chip Exports (35,000 GB300-equivalent each)", "19 Nov 2025", "https://www.commerce.gov/news/press-releases/2025/11/statement-uae-and-saudi-chip-exports"),
    ("Reuters, US authorizes export of advanced American semiconductors to companies in Saudi, UAE", "19 Nov 2025", "https://www.reuters.com/business/us-authorizes-export-advanced-american-semiconductors-companies-saudi-uae-2025-11-19/"),
    ("OpenAI, Introducing Stargate UAE (1 GW, 200 MW in 2026)", "May 2025", "https://openai.com/index/introducing-stargate-uae/"),
    ("G42, Global Tech Alliance Launches Stargate UAE", "May 2025", "https://www.g42.ai/resources/news/global-tech-alliance-launches-stargate-uae"),
    ("The Tech Capital, First 200 MW of Stargate UAE will come online in 2026 says G42 chief (5 GW campus)", "15 Oct 2025", "https://thetechcapital.com/first-200mw-of-stargate-uae-will-come-online-in-2026-says-g42-chief"),
    ("European Commission, AI Factories policy page (19 AI Factories, up to seven gigafactories, EUR 10B)", "12 Aug 2026", "https://digital-strategy.ec.europa.eu/en/policies/ai-factories"),
    ("EIB Group and European Commission join forces to finance AI gigafactories (~100,000 chips each)", "2025", "https://www.eib.org/en/press/all/2025-491-eib-group-and-european-commission-join-forces-to-finance-ai-gigafactories"),
    ("Augur (Mulani, Fitz, Fort), Who Owns AI? (sovereign funds and programmes, data to May 2026)", "May 2026", "https://whoownsai.org/who-owns-ai.pdf"),
    ("Deutsche Telekom, For a sovereign Germany: Deutsche Telekom launches Industrial AI Cloud with NVIDIA (10,000 Blackwell GPUs, Munich)", "4 Nov 2025", "https://www.telekom.com/en/newsroom/latest-updates/media-information/2025/11/for-a-sovereign-germany"),
    ("Nscale, UK AI infrastructure commitment with Microsoft, NVIDIA and OpenAI (58,640 GPUs; Loughton 23,040 GB300)", "16 Sep 2025", "https://www.nscale.com/press-releases/nscale-uk-ai-infrastructure-announcement"),
    ("OpenAI, Introducing Stargate UK", "Sep 2025", "https://openai.com/index/introducing-stargate-uk/"),
    ("NVIDIA, NVIDIA, South Korea Government and Industrial Giants Build AI Infrastructure (260,000 GPUs)", "31 Oct 2025", "https://investor.nvidia.com/news/press-release-details/2025/NVIDIA-South-Korea-Government-and-Industrial-Giants-Build-AI-Infrastructure-and-Ecosystem-to-Fuel-Korea-Innovation-Industries-and-Jobs/default.aspx"),
    ("Spheron, GPU Cloud Providers in Japan 2026 (27,500 Rubin, METI JPY 72.5B, Sakura, prices)", "Oct 2026", "https://www.spheron.network/blog/gpu-cloud-providers-japan-2026/"),
    ("IndiaAI, Compute Capacity (38,000 GPUs, INR 150 per hour)", "2026", "https://indiaai.gov.in/hub/indiaai-compute-capacity"),
    ("Introl, Malaysia and Thailand: Emerging AI Data Center Hubs (YTL 600 MW GB200, budgets)", "Dec 2025", "https://introl.com/blog/malaysia-thailand-emerging-ai-data-center-hubs-southeast-asia"),
    ("Digital in Asia, Who is Building AI Data Centres in Southeast Asia in 2026", "2026", "https://digitalinasia.com/southeast-asia-ai-data-centre-boom/"),
    ("Google Cloud, Launches New Cloud Region in Thailand ($1B; THB 1.4T value)", "21 Jan 2026", "https://www.googlecloudpresscorner.com/2026-01-21-Google-Cloud-Launches-New-Cloud-Region-in-Thailand,-Bolstering-its-Commitment-to-Advancing-the-Countrys-AI-Driven-Digital-Economy"),
    ("MarketsandMarkets, Sovereign AI Market to Reach USD 148.0 Billion by 2032", "19 Aug 2026", "https://finance.yahoo.com/technology/ai/articles/sovereign-ai-market-reach-usd-151500186.html"),
    ("Grand View Research, Sovereign AI Market (2026-2033)", "2026", "https://www.grandviewresearch.com/industry-analysis/sovereign-ai-market-report"),
    ("Precedence Research, Sovereign AI Infrastructure Market", "2026", "https://www.precedenceresearch.com/sovereign-ai-infrastructure-market"),
    ("SNS Insider, Sovereign AI Infrastructure Market", "2026", "https://www.snsinsider.com/reports/sovereign-ai-infrastructure-market-10918"),
    ("Gartner, Worldwide AI Spending to Grow 49.5% in 2026", "16 Sep 2026", "https://www.gartner.com/en/newsroom/press-releases/2026-09-16-gartner-forecasts-worldwide-ai-spending-to-grow-49-point-5-percent-in-2026"),
    ("Gartner, Over $3.5 Trillion of AI Infrastructure Spending Could Be Exposed to Impairment by End of 2030", "14 Oct 2026", "https://www.gartner.com/en/newsroom/press-releases/2026-10-14-gartner-says-over-us-dollars-three-and-a-half-trillion-of-artificial-intelligence-infrastructure-spending-could-be-exposed-to-impairment-by-the-end-of-2030"),
    ("Epoch AI, Total cost of ownership of a one-gigawatt AI data center ($38B capex, $8.5B/yr)", "14-19 May 2026", "https://epoch.ai/data-insights/ai-datacenter-cost-breakdown"),
    ("Carnegie Endowment, The Compute Coalition: How to Build the Future of AI in the Free World", "Jun 2026", "https://carnegieendowment.org/research/2026/06/the-compute-coalition-how-to-build-the-future-of-ai-in-the-free-world"),
    ("IDCA, Global Data Center Report 2026 (67.7 GW footprint; US 29.2 GW; Germany 5.5 GW)", "2026", "https://www.idc-a.org/insights/Sk2PFZCost98Eqhk2FUA"),
    ("Brookings, Global energy demands within the AI regulatory landscape", "2026", "https://www.brookings.edu/articles/global-energy-demands-within-the-ai-regulatory-landscape/"),
    ("Qovery, European GPU Clouds: 8 Alternatives to AWS and Azure (Sep 2026 list prices)", "Sep 2026", "https://www.qovery.com/blog/european-gpu-cloud-alternatives-aws-azure-comparison"),
    ("Spheron, OVHcloud GPU Pricing 2026: H100 Cost and Sovereign AI (EUR 180M Commission contract)", "Oct 2026", "https://www.spheron.network/blog/ovhcloud-gpu-pricing-2026-h100-cost-and-sovereign-ai/"),
    ("computeprices.com, Nebius GPU pricing (B200 $8.50, H100 $4.50)", "14 Oct 2026", "https://computeprices.com/providers/nebius"),
    ("EIA, Electricity Monthly Update, End Use, July 2026 (industrial 9.77c/kWh)", "Jul 2026", "https://www.eia.gov/electricity/monthly/update/end-use.php"),
    ("NVIDIA, Financial Results for Second Quarter Fiscal 2027 ($89.0B Data Center, 75% GM)", "Aug 2026", "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-second-quarter-fiscal-2027"),
    ("Jensen Huang, NVIDIA AI Factory Compute Is Becoming an Investable Asset Class (H100 index $2.70)", "12 Aug 2026", "https://blogs.nvidia.com/blog/nvidia-ai-factory-compute/"),
    ("PwC, Sizing the Prize via World Economic Forum ($15.7T, +14% GDP by 2030)", "2017", "https://www.weforum.org/stories/geographies-in-depth/the-global-economy-will-be-14-bigger-in-2030-because-of-ai"),
    ("IDC, AI to drive $19.9 trillion through 2030", "2024", "https://my.idc.com/getdoc.jsp?containerId=prUS52600524"),
    ("Bruegel, How can Europe address its pressing AI compute infrastructure shortfall?", "Sep 2026", "https://www.bruegel.org/policy-brief/how-can-europe-address-its-pressing-ai-compute-infrastructure-shortfall"),
    ("Futurum, AI Capex 2026: The $690B Infrastructure Sprint (13 EuroHPC AI factories; European AI server spend $47B)", "12 Feb 2026", "https://futurumgroup.com/insights/ai-capex-2026-the-690b-infrastructure-sprint/"),
    ("Reuters, Nvidia's China AI-chip share falls as Huawei closes gap (IDC: 34% vs 28%, H1 2026)", "13 Oct 2026", "https://www.reuters.com/business/retail-consumer/nvidias-china-ai-chip-share-falls-huawei-closes-gap-report-finds-2026-10-13/"),
    ("Avid Solutions summarising McKinsey ($5.2T for 156 GW by 2030) and JLL ($11.3M/MW construction 2026)", "Jan 2026", "https://avidsolutionsinc.com/13-data-center-growth-projections-that-will-shape-2026-2030/"),
]
sr = 5
for i, (s, d, u) in enumerate(SOURCES, start=1):
    S.cell(row=sr, column=1, value=i).font = BLACK; S.cell(row=sr, column=2, value=s).font = BLACK; S.cell(row=sr, column=3, value=d).font = BLACK
    c = S.cell(row=sr, column=4, value=u); c.font = Font(name=F, size=9, color="0F6B63", underline="single"); c.hyperlink = u; sr += 1
S.freeze_panes = "A5"

wb.save(OUT)
import json
json.dump({"BU": BU, "GT": GT, "KC": KC, "PR": PR, "VV": VV, "qstart": qstart, "first": first, "last": last, "tot": tot}, open("/home/claude/ai-radar/research/sovereign-ai/model_refs.json", "w"))
print("saved", OUT, "assumptions:", row - 5)
