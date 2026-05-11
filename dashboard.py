#!/usr/bin/env python3
"""
LLG Industries — Sales Dashboard Generator
Data: 2025 onwards only (pre-2025 margin data unreliable)
"""

import sys
import re as _re
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
    import plotly.graph_objects as go
except ImportError:
    print("Run:  pip install pandas openpyxl plotly")
    sys.exit(1)

OUTPUT_DIR = Path(__file__).parent.parent / "03_Dashboard"

EXCLUDE_PREFIXES = {
    "5 CENTS", "A3 POSTER", "BRB SILICONE", "ELECTRICITY", "EMPTY DRUM",
    "FLS LIGHTER", "HAMPER", "PURE SILICONE", "RENTAL", "RUISIL",
    "SAMPLE", "SHELL", "SILICONE", "TEA TREE OIL", "TERGITOL",
    "10 LITRE JAR", "TRANSPORTATION", "XE55",
}

BRAND_MAP = {
    # ── Licin Licin (Own Brand) ────────────────────────────────────────────
    "LICIN LICIN":      "Licin Licin",
    "FLOOR CLEANER":    "Licin Licin",
    "TOILET KING":      "Licin Licin",
    "LEMON GRASS":      "Licin Licin",
    "LLG LICIN":        "Licin Licin",
    "LLG FLOOR":        "Licin Licin",
    "LLG TOILET":       "Licin Licin",
    "ROSE LICIN":       "Licin Licin",
    "GREEN TOILET":     "Licin Licin",
    "FLOOR DETERGENT":  "Licin Licin",
    "DISH WASHER":      "Licin Licin",
    "DISHWASHING":      "Licin Licin",
    "DISH LIQUID":      "Licin Licin",
    # ── Other Own Brands ──────────────────────────────────────────────────
    "L LADY":           "L Lady",
    "HATA":             "HATA",
    "CNC":              "CNC",
    "RADIX":            "Radix",
    "MOSAKO":           "Mosako",
    "PET REJOICE":      "Pet Rejoice",
    # ── OEM Brands ────────────────────────────────────────────────────────
    "JEEMAT":           "Jeemat (OEM)",
    "KYLAT":            "Kylat (OEM)",
    "PET WIPE":         "Pet Wipe (OEM)",
    "AASION":           "Aasion (OEM)",
    "DURO":             "Duro (OEM)",
    "HILO":             "Hilo (OEM)",
    "BIO CLEANSE":      "Bio Cleanse (OEM)",
    "PROSPERITY":       "Prosperity (OEM)",
    "HAN'S PET":        "Han's Pet (OEM)",
    "HANZ":             "Han's Pet (OEM)",
    "DEPEX":            "Depex (OEM)",
    "ECENTRA":          "Ecentra (OEM)",
    "MAGIC 7":          "Magic 7 (OEM)",
    "PET SHAMPOO":      "Pet Shampoo (OEM)",
    "WET TISSUE":       "Wet Tissue (OEM)",
    "MAXICARE":         "Maxicare (OEM)",
    # ── Industrial Range ──────────────────────────────────────────────────
    "801PRO":           "Industrial Range",
    "802PRO":           "Industrial Range",
    "103PRO":           "Industrial Range",
    "201PRO":           "Industrial Range",
    "302PRO":           "Industrial Range",
    "CITRONELLA":       "Industrial Range",
    "SNOW CAR WASH":    "Industrial Range",
    "TYRE SHINE":       "Industrial Range",
    "QUICK WAX":        "Industrial Range",
}
COLORS = {
    "Licin Licin":       "#1a3c5e",
    "HATA":              "#27ae60",
    "CNC":               "#e67e22",
    "L Lady":            "#8e44ad",
    "Jeemat (OEM)":      "#2980b9",
    "Kylat (OEM)":       "#16a085",
    "Pet Wipe (OEM)":    "#3498db",
    "Aasion (OEM)":      "#1abc9c",
    "Duro (OEM)":        "#e74c3c",
    "Hilo (OEM)":        "#f1c40f",
    "Bio Cleanse (OEM)": "#9b59b6",
    "Prosperity (OEM)":  "#e91e63",
    "Han's Pet (OEM)":   "#795548",
    "Radix":             "#c0392b",
    "Mosako":            "#7f8c8d",
    "Pet Rejoice":       "#f39c12",
    "Depex (OEM)":       "#d35400",
    "Ecentra (OEM)":     "#2c3e50",
    "Magic 7 (OEM)":     "#00bcd4",
    "Pet Shampoo (OEM)": "#ff7043",
    "Wet Tissue (OEM)":  "#ab47bc",
    "Maxicare (OEM)":    "#26a69a",
    "Industrial Range":  "#78909c",
    "Other":             "#bdc3c7",
    "LLE":               "#2980b9",
    "LLG":               "#8e44ad",
    "AGENT1":            "#e74c3c",
    "AGENT2":            "#27ae60",
    "AGENT3":            "#8e44ad",
    "New":               "#27ae60",
    "Returning":         "#2980b9",
}
BG      = "#f8f9fa"
CARD_BG = "#ffffff"

ECOMM_KEYWORDS = {"SHOPEE", "LAZADA", "TIKTOK", "ECOMMERCE", "E-COMMERCE", "E COMMERCE"}

def _is_ecommerce(customer: str) -> bool:
    cu = str(customer).upper()
    return any(kw in cu for kw in ECOMM_KEYWORDS)

def _norm_credit(val):
    if pd.isna(val):
        return None
    s = str(val).strip().upper().replace(".", "").replace(" ", "").replace("*", "")
    if s in ("CASH", "COD"):
        return "Cash / C.O.D"
    if s in ("30P30", "30C30", "30R30", "30DAYS", "30"):
        return "30 Days"
    m = _re.match(r"30[A-Z](\d+)$", s)
    if m:
        return f"{m.group(1)} Days"
    m = _re.match(r"(\d+)DAYS?$", s)
    if m:
        return f"{m.group(1)} Days"
    m = _re.match(r"(\d+)$", s)
    if m:
        return f"{m.group(1)} Days"
    return str(val).strip()

def _days_key(label):
    if label == "Cash / C.O.D": return 0
    if label == "Consignment":   return 9999
    m = _re.match(r"(\d+)", label)
    return int(m.group(1)) if m else 9998

# ── Load Data ──────────────────────────────────────────────────────────────────

def _parse_date(series):
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, origin="1899-12-30", unit="D", errors="coerce")
    return pd.to_datetime(series, errors="coerce")

def load(filepath: str, company: str) -> pd.DataFrame:
    df = pd.read_excel(filepath)
    df["Date"] = _parse_date(df["Post Date"])
    df = df[df["Date"].dt.year >= 2025]
    df["YearMonth"]  = df["Date"].dt.to_period("M").astype(str)
    df["MonthLabel"] = df["Date"].dt.strftime("%b %Y")
    df["Year"]       = df["Date"].dt.year
    df["Month"]      = df["Date"].dt.month
    df["MonthName"]  = df["Date"].dt.strftime("%b")
    df["Brand"] = df["Item Desc"].apply(
        lambda x: next((v for k, v in BRAND_MAP.items() if str(x).strip().upper().startswith(k)), "Other")
    )
    df = df[~df["Item Desc"].astype(str).str.strip().str.upper().apply(
        lambda x: any(x.startswith(p) for p in EXCLUDE_PREFIXES)
    )]
    df["Company"] = company
    df = df.rename(columns={
        "Local Item Sales": "Revenue_RM",
        "Item Cost":        "COGS_RM",
        "Item P/L":         "Margin_RM",
        "Item Margin":      "Margin_Pct",
        "CompanyCategory":  "Channel",
        "COMPANYNAME":      "Customer",
        "Item Desc":        "Product",
    })
    df = df[df["Agent"].astype(str).str.upper() != "LIC"]
    return df[df["Revenue_RM"].notna() & (df["Revenue_RM"] != 0) & df["Product"].notna()]


def load_foc(filepath: str, company: str) -> pd.DataFrame:
    df = pd.read_excel(filepath)
    df["Date"] = _parse_date(df["Post Date"])
    df = df[
        (df["Date"].dt.year >= 2025) &
        df["Item Desc"].notna() &
        (~df["Agent"].astype(str).str.upper().eq("LIC"))
    ]
    foc = df[(df["Local Item Sales"].fillna(0) == 0) & (df["Item Cost"].fillna(0) > 0)].copy()
    foc["Company"] = company
    foc = foc.rename(columns={"Item Desc": "Product", "Item Cost": "Cost_RM", "COMPANYNAME": "Customer"})
    return foc[["Date", "Product", "Cost_RM", "Customer", "Company"]]


def load_full(filepath: str) -> pd.DataFrame:
    """Load all years without date filter — for customer history analysis only."""
    df = pd.read_excel(filepath)
    df["Date"] = _parse_date(df["Post Date"])
    df["Year"] = df["Date"].dt.year.astype("Int64")
    df = df.rename(columns={"Local Item Sales": "Revenue_RM", "COMPANYNAME": "Customer"})
    df = df[df["Agent"].astype(str).str.upper() != "LIC"]
    df = df[df["Revenue_RM"].notna() & (df["Revenue_RM"] != 0) & df["Customer"].notna() & df["Date"].notna()]
    return df[["Date", "Year", "Customer", "Revenue_RM", "Agent"]]

# ── Helpers ────────────────────────────────────────────────────────────────────

def fmt_rm(val):
    if abs(val) >= 1_000_000:
        return f"RM {val/1_000_000:.2f}M"
    if abs(val) >= 1_000:
        return f"RM {val/1_000:.1f}K"
    return f"RM {val:.2f}"

def card_html(label, value, sub="", color="#1a3c5e"):
    return f"""
    <div style="background:{CARD_BG};border-radius:12px;padding:24px 20px;
                box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;flex:1;min-width:160px">
      <div style="color:#888;font-size:13px;font-weight:600;text-transform:uppercase;
                  letter-spacing:0.5px;margin-bottom:8px">{label}</div>
      <div style="color:{color};font-size:28px;font-weight:700;line-height:1">{value}</div>
      {f'<div style="color:#aaa;font-size:12px;margin-top:6px">{sub}</div>' if sub else ''}
    </div>"""

def section_header(title):
    return f"""
    <div style="font-size:16px;font-weight:700;color:#1a3c5e;
                padding:16px 4px 6px;border-bottom:2px solid #1a3c5e;
                margin-bottom:12px">{title}</div>"""

def fig_html(fig):
    return fig.to_html(full_html=False, include_plotlyjs=False,
                       config={"displayModeBar": False, "responsive": True})

def rm_axis(fig, axis="yaxis"):
    fig.update_layout(**{axis: dict(tickprefix="RM ", tickformat=",.0f")})

# ── Build Dashboard ────────────────────────────────────────────────────────────

def build_dashboard(df: pd.DataFrame, foc_cost: float, df_full: pd.DataFrame, report_title: str = "LLG Industries Sdn Bhd — Sales Dashboard") -> str:
    # BrandType, SalesChannel and type colors — computed once, used throughout
    df = df.copy()
    df["BrandType"] = df["Brand"].apply(
        lambda b: "OEM" if "(OEM)" in str(b)
        else ("Industrial Range" if b == "Industrial Range" else "Own Brand")
    )
    df["SalesChannel"] = df["Customer"].apply(
        lambda c: "Ecommerce" if _is_ecommerce(c) else "Traditional"
    )
    OWN_COLOR   = "#1a3c5e"
    OEM_COLOR   = "#2980b9"
    IND_COLOR   = "#78909c"
    ECOMM_COLOR = "#8e44ad"
    TRAD_COLOR  = "#27ae60"

    rev          = df["Revenue_RM"].sum()
    mgn          = df["Margin_RM"].sum()
    mgp          = mgn / rev * 100 if rev else 0
    unique_docs  = df["Doc No"].nunique()
    n_months     = df["YearMonth"].nunique()
    period_start = df["Date"].min().strftime("%b %Y")
    period_end   = df["Date"].max().strftime("%b %Y")

    # ── Summary Cards ─────────────────────────────────────────────────────────
    cards = "".join([
        card_html("Total Revenue",        fmt_rm(rev),              f"{period_start} – {period_end}", "#1a3c5e"),
        card_html("Total Gross Margin",   fmt_rm(mgn),              f"{mgp:.1f}% margin",             "#27ae60"),
        card_html("Avg Monthly Revenue",  fmt_rm(rev / n_months),   "per month",                      "#2980b9"),
        card_html("Invoices (Doc No)",    f"{unique_docs:,}",       "unique documents",               "#8e44ad"),
        card_html("Marketing Cost (FOC)", fmt_rm(foc_cost),         "free goods at cost",             "#e67e22"),
    ])

    # ── YTD Comparison ────────────────────────────────────────────────────────
    MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    latest_date = df["Date"].max()
    ytd_year    = latest_date.year
    ytd_month   = latest_date.month
    ytd_lbl     = f"Jan – {MONTHS_SHORT[ytd_month-1]} {ytd_year}"
    prev_lbl    = f"Jan – {MONTHS_SHORT[ytd_month-1]} {ytd_year-1}"

    ytd_cur  = df[(df["Year"] == ytd_year)   & (df["Month"] <= ytd_month)]
    ytd_prev = df[(df["Year"] == ytd_year-1) & (df["Month"] <= ytd_month)]

    ytd_cur_rev   = ytd_cur["Revenue_RM"].sum()
    ytd_prev_rev  = ytd_prev["Revenue_RM"].sum()
    ytd_cur_mgn   = ytd_cur["Margin_RM"].sum()
    ytd_prev_mgn  = ytd_prev["Margin_RM"].sum()
    ytd_cur_mgp   = ytd_cur_mgn  / ytd_cur_rev  * 100 if ytd_cur_rev  else 0
    ytd_prev_mgp  = ytd_prev_mgn / ytd_prev_rev * 100 if ytd_prev_rev else 0

    ytd_rev_chg = ytd_cur_rev - ytd_prev_rev
    ytd_rev_pct = ytd_rev_chg / ytd_prev_rev * 100 if ytd_prev_rev else 0
    ytd_mgn_chg = ytd_cur_mgn - ytd_prev_mgn
    ytd_mgn_pct = ytd_mgn_chg / ytd_prev_mgn * 100 if ytd_prev_mgn else 0

    def _chg_html(val, pct):
        arrow = "▲" if val >= 0 else "▼"
        clr   = "#27ae60" if val >= 0 else "#e74c3c"
        return f'<span style="color:{clr}">{arrow} {fmt_rm(abs(val))} ({abs(pct):.1f}%)</span>'

    ytd_cards = "".join([
        card_html(f"YTD {ytd_year} Revenue",        fmt_rm(ytd_cur_rev),  ytd_lbl,                          "#1a3c5e"),
        card_html(f"YTD {ytd_year-1} Revenue",      fmt_rm(ytd_prev_rev), prev_lbl,                         "#2980b9"),
        card_html("YoY Revenue Change",              _chg_html(ytd_rev_chg, ytd_rev_pct), "vs same period",  "#1a3c5e"),
        card_html(f"YTD {ytd_year} Gross Margin",   fmt_rm(ytd_cur_mgn),  f"{ytd_cur_mgp:.1f}% margin",     "#27ae60"),
        card_html(f"YTD {ytd_year-1} Gross Margin", fmt_rm(ytd_prev_mgn), f"{ytd_prev_mgp:.1f}% margin",    "#16a085"),
        card_html("YoY Margin Change",               _chg_html(ytd_mgn_chg, ytd_mgn_pct), "vs same period", "#1a3c5e"),
    ])

    fig_ytd = go.Figure()
    for year, color in [(ytd_year-1, "#2980b9"), (ytd_year, "#e67e22")]:
        by_m = df[(df["Year"] == year) & (df["Month"] <= ytd_month)].groupby("Month")["Revenue_RM"].sum()
        fig_ytd.add_trace(go.Bar(
            x=[MONTHS_SHORT[m-1] for m in range(1, ytd_month+1)],
            y=[float(by_m.get(m, 0)) for m in range(1, ytd_month+1)],
            name=str(year), marker_color=color,
            hovertemplate="%{x} " + str(year) + "<br>RM %{y:,.0f}<extra></extra>",
        ))
    fig_ytd.update_layout(
        title=f"YTD Monthly Revenue — {ytd_year-1} vs {ytd_year} ({ytd_lbl.replace(str(ytd_year),'')})",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(type="category"),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=40, l=90, r=20), height=340,
    )

    # ── Monthly Trend ─────────────────────────────────────────────────────────
    monthly = df.groupby(["YearMonth","MonthLabel","Company"]).agg(
        Revenue=("Revenue_RM","sum")
    ).reset_index().sort_values("YearMonth")

    fig_trend = go.Figure()
    for company, color in [("LLE", COLORS["LLE"]), ("LLG", COLORS["LLG"])]:
        m = monthly[monthly["Company"] == company]
        fig_trend.add_trace(go.Scatter(
            x=m["MonthLabel"].tolist(), y=m["Revenue"].tolist(),
            name=company, mode="lines+markers",
            line=dict(color=color, width=2.5), marker=dict(size=5),
            hovertemplate="%{x}<br>Revenue: RM %{y:,.0f}<extra>" + company + "</extra>",
        ))
    fig_trend.update_layout(
        title="Monthly Revenue — LLE vs LLG",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=80, l=90, r=20), height=340
    )

    # ── 2025 vs 2026 ──────────────────────────────────────────────────────────
    MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    fig_yoy = go.Figure()
    for year, color in [(2025, "#2980b9"), (2026, "#e67e22")]:
        by_m = df[df["Year"] == year].groupby("Month")["Revenue_RM"].sum()
        fig_yoy.add_trace(go.Bar(
            x=MONTHS,
            y=[float(by_m.get(m, 0)) for m in range(1, 13)],
            name=str(year), marker_color=color,
            hovertemplate="%{x} " + str(year) + "<br>RM %{y:,.0f}<extra></extra>",
        ))
    fig_yoy.update_layout(
        title="Monthly Sales — 2025 vs 2026",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(type="category"),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=40, l=90, r=20), height=340
    )

    # ── Brand ─────────────────────────────────────────────────────────────────
    brand = df.groupby("Brand").agg(
        Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum")
    ).reset_index()
    brand["MarginPct"] = (brand["Margin"] / brand["Revenue"] * 100).round(1)
    brand = brand[brand["Revenue"] > 0].sort_values("Revenue", ascending=True)

    fig_brand_rev = go.Figure()
    for _, r in brand.iterrows():
        fig_brand_rev.add_trace(go.Bar(
            x=[r["Revenue"]], y=[r["Brand"]], orientation="h",
            marker_color=COLORS.get(r["Brand"], "#95a5a6"), showlegend=False,
            hovertemplate=f"{r['Brand']}<br>RM %{{x:,.0f}}<extra></extra>"
        ))
    fig_brand_rev.update_layout(
        title=f"Revenue by Brand — Total",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=20, l=140, r=20), height=420
    )

    fig_brand_mgn = go.Figure()
    for _, r in brand.iterrows():
        fig_brand_mgn.add_trace(go.Bar(
            x=[r["MarginPct"]], y=[r["Brand"]], orientation="h",
            marker_color=COLORS.get(r["Brand"], "#95a5a6"), showlegend=False,
            hovertemplate=f"{r['Brand']}<br>Margin: %{{x:.1f}}%<extra></extra>"
        ))
    fig_brand_mgn.update_layout(
        title="Gross Margin % by Brand",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(ticksuffix="%", tickformat=".1f"),
        margin=dict(t=50, b=20, l=140, r=20), height=420
    )

    # ── Channel ───────────────────────────────────────────────────────────────
    channel = df.groupby("Channel").agg(Revenue=("Revenue_RM","sum")).reset_index()
    channel = channel.sort_values("Revenue", ascending=False).head(10).sort_values("Revenue", ascending=True)

    fig_channel = go.Figure(go.Bar(
        x=channel["Revenue"].tolist(), y=channel["Channel"].tolist(), orientation="h",
        marker_color="#1a3c5e",
        hovertemplate="%{y}<br>Total: RM %{x:,.0f}<extra></extra>",
    ))
    fig_channel.update_layout(
        title=f"Total Revenue by Channel ({period_start}–{period_end})",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=20, l=120, r=20), height=340
    )

    # ── Top 12 Customers ──────────────────────────────────────────────────────
    cust = df.groupby("Customer")["Revenue_RM"].sum().nlargest(12).reset_index()
    cust = cust.sort_values("Revenue_RM", ascending=True)

    fig_cust = go.Figure(go.Bar(
        x=cust["Revenue_RM"].tolist(), y=cust["Customer"].tolist(), orientation="h",
        marker_color="#2980b9",
        hovertemplate="%{y}<br>Total: RM %{x:,.0f}<extra></extra>",
    ))
    fig_cust.update_layout(
        title=f"Top 12 Customers — Total Revenue ({period_start}–{period_end})",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=20, l=220, r=20), height=380
    )

    # ── Top Products Table (split Own Brand / OEM) ────────────────────────────
    prod_full = df.groupby(["Product","Brand","BrandType"]).agg(
        Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum")
    ).reset_index()
    prod_full["AvgMonthly"] = prod_full["Revenue"] / n_months
    prod_full["MarginPct"]  = (prod_full["Margin"] / prod_full["Revenue"] * 100).round(1)

    def _prod_section(btype, hcolor, n=15):
        sub = (prod_full[prod_full["BrandType"] == btype]
               .sort_values("Revenue", ascending=False).head(n))
        sub_total = sub["Revenue"].sum()
        rows = ""
        for i, (_, r) in enumerate(sub.iterrows(), 1):
            bg = "#f8f9fa" if i % 2 == 0 else CARD_BG
            mgn_color = "#27ae60" if r["MarginPct"] >= 45 else "#e74c3c"
            rows += (
                f'<tr style="background:{bg}">'
                f'<td style="padding:7px 10px;text-align:center;color:#888">{i}</td>'
                f'<td style="padding:7px 10px">{r["Product"]}</td>'
                f'<td style="padding:7px 10px;text-align:right;font-weight:600">RM {r["Revenue"]:,.0f}</td>'
                f'<td style="padding:7px 10px;text-align:right">RM {r["AvgMonthly"]:,.0f}</td>'
                f'<td style="padding:7px 10px;text-align:right'
                f';color:{mgn_color};font-weight:600">{r["MarginPct"]:.1f}%</td>'
                f'<td style="padding:7px 10px;text-align:right">'
                f'{r["Revenue"]/sub_total*100:.1f}%</td>'
                f'</tr>'
            )
        return f"""
        <div style="margin-bottom:24px">
          <div style="font-size:14px;font-weight:700;color:{hcolor};margin-bottom:8px;
                      padding-bottom:4px;border-bottom:2px solid {hcolor}">
            {btype} — Top {n} Products
          </div>
          <table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead><tr style="background:{hcolor};color:white">
              <th style="padding:7px 10px">#</th>
              <th style="padding:7px 10px;text-align:left">Product</th>
              <th style="padding:7px 10px;text-align:right">Revenue</th>
              <th style="padding:7px 10px;text-align:right">Avg/Month</th>
              <th style="padding:7px 10px;text-align:right">Margin %</th>
              <th style="padding:7px 10px;text-align:right">% within type</th>
            </tr></thead>
            <tbody>{rows}</tbody>
          </table>
        </div>"""

    prod_table = f"""
    <div style="padding:8px 16px">
      <div style="font-size:15px;font-weight:700;color:#333;margin-bottom:16px">
        Top Products by Brand Type — {period_start}–{period_end}
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:24px">
        {_prod_section("Own Brand", OWN_COLOR)}
        {_prod_section("OEM", OEM_COLOR)}
      </div>
      {_prod_section("Industrial Range", IND_COLOR, n=10)}
    </div>"""

    # ── Agent Analysis ────────────────────────────────────────────────────────
    agents_df = df[df["Agent"].isin(["AGENT1", "AGENT2"])].copy()

    agent_monthly = agents_df.groupby(["YearMonth","MonthLabel","Agent"]).agg(
        Revenue=("Revenue_RM","sum")
    ).reset_index().sort_values("YearMonth")

    fig_agent = go.Figure()
    for agent in ["AGENT1", "AGENT2"]:
        am = agent_monthly[agent_monthly["Agent"] == agent]
        fig_agent.add_trace(go.Bar(
            x=am["MonthLabel"].tolist(), y=am["Revenue"].tolist(),
            name=agent, marker_color=COLORS.get(agent, "#333"),
            hovertemplate="%{x}<br>" + agent + ": RM %{y:,.0f}<extra></extra>",
        ))
    fig_agent.update_layout(
        title="In-House Sales — Monthly Revenue by Agent",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=80, l=90, r=20), height=340
    )

    agent_cust_df = agents_df.groupby(["Agent","Customer"])["Revenue_RM"].sum().reset_index()
    fig_agent_cust = go.Figure()
    for agent in ["AGENT1", "AGENT2"]:
        ac = agent_cust_df[agent_cust_df["Agent"] == agent].nlargest(8, "Revenue_RM").sort_values("Revenue_RM", ascending=True)
        fig_agent_cust.add_trace(go.Bar(
            x=ac["Revenue_RM"].tolist(), y=ac["Customer"].tolist(), orientation="h",
            name=agent, marker_color=COLORS.get(agent, "#333"),
            hovertemplate="%{y}<br>" + agent + ": RM %{x:,.0f}<extra></extra>",
        ))
    fig_agent_cust.update_layout(
        title="Top Customers by Agent — Total Revenue",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=20, l=220, r=20), height=400
    )

    agent_sum = agents_df.groupby("Agent").agg(
        Revenue=("Revenue_RM","sum"),
        Margin=("Margin_RM","sum"),
        Invoices=("Doc No","nunique"),
        Customers=("Customer","nunique"),
        Products=("Product","nunique"),
    ).reset_index()
    agent_sum["MarginPct"]   = (agent_sum["Margin"] / agent_sum["Revenue"] * 100)
    agent_sum["AvgOrder"]    = agent_sum["Revenue"] / agent_sum["Invoices"]
    agent_sum["AvgMonthly"]  = agent_sum["Revenue"] / n_months
    agent_sum["RevShare"]    = agent_sum["Revenue"] / rev * 100

    agent_rows = ""
    for _, r in agent_sum.iterrows():
        agent_rows += (
            f'<tr>'
            f'<td style="padding:9px 12px;font-weight:700;color:{COLORS.get(r["Agent"],"#333")}">{r["Agent"]}</td>'
            f'<td style="padding:9px 12px;text-align:right;font-weight:600">RM {r["Revenue"]:,.0f}</td>'
            f'<td style="padding:9px 12px;text-align:right">RM {r["AvgMonthly"]:,.0f}</td>'
            f'<td style="padding:9px 12px;text-align:right">{r["RevShare"]:.1f}%</td>'
            f'<td style="padding:9px 12px;text-align:right">{r["MarginPct"]:.1f}%</td>'
            f'<td style="padding:9px 12px;text-align:right">{int(r["Invoices"]):,}</td>'
            f'<td style="padding:9px 12px;text-align:right">RM {r["AvgOrder"]:,.0f}</td>'
            f'<td style="padding:9px 12px;text-align:right">{int(r["Customers"])}</td>'
            f'<td style="padding:9px 12px;text-align:right" title="Unique SKUs sold">{int(r["Products"])}</td>'
            f'</tr>'
        )

    agent_table = f"""
    <div style="padding:8px 16px 16px">
      <div style="font-size:15px;font-weight:700;color:#333;margin-bottom:12px">
        Agent Summary ({period_start}–{period_end})
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:#1a3c5e;color:white">
            <th style="padding:8px 12px;text-align:left">Agent</th>
            <th style="padding:8px 12px;text-align:right">Total Revenue</th>
            <th style="padding:8px 12px;text-align:right">Avg Monthly</th>
            <th style="padding:8px 12px;text-align:right">% of Total</th>
            <th style="padding:8px 12px;text-align:right">Margin %</th>
            <th style="padding:8px 12px;text-align:right">Invoices</th>
            <th style="padding:8px 12px;text-align:right">Avg Invoice</th>
            <th style="padding:8px 12px;text-align:right">Customers</th>
            <th style="padding:8px 12px;text-align:right">SKUs Sold</th>
          </tr>
        </thead>
        <tbody>{agent_rows}</tbody>
      </table>
      <div style="font-size:11px;color:#aaa;margin-top:6px">
        * SKUs Sold = number of distinct product lines handled by this agent
      </div>
    </div>"""

    # ── 1. Customer Repeat Rate ───────────────────────────────────────────────
    # Use full history to determine each customer's first-ever month
    df_full_ym = df_full.copy()
    df_full_ym["YearMonth"] = df_full_ym["Date"].dt.to_period("M").astype(str)
    first_month_ever = df_full_ym.groupby("Customer")["YearMonth"].min().rename("FirstEverMonth")

    cust_monthly = df.groupby(["YearMonth","MonthLabel","Customer"]).size().reset_index(name="rows")
    cust_monthly = cust_monthly.merge(first_month_ever, on="Customer", how="left")
    cust_monthly["Type"] = cust_monthly.apply(
        lambda r: "New" if r["YearMonth"] == r["FirstEverMonth"] else "Returning", axis=1
    )
    repeat_chart = cust_monthly.groupby(["YearMonth","MonthLabel","Type"])["Customer"].nunique().reset_index()
    repeat_chart = repeat_chart.sort_values("YearMonth")

    # Fix x-axis order: use sorted unique months so all traces share the same order
    sorted_months = repeat_chart.drop_duplicates("YearMonth").sort_values("YearMonth")[["YearMonth","MonthLabel"]].values.tolist()
    month_order   = [m[1] for m in sorted_months]

    fig_repeat = go.Figure()
    for ctype in ["Returning", "New"]:
        rc = repeat_chart[repeat_chart["Type"] == ctype].set_index("MonthLabel")
        fig_repeat.add_trace(go.Bar(
            x=month_order,
            y=[rc.loc[m, "Customer"] if m in rc.index else 0 for m in month_order],
            name=ctype, marker_color=COLORS.get(ctype, "#333"),
            hovertemplate="%{x}<br>" + ctype + " customers: %{y}<extra></extra>",
        ))
    fig_repeat.update_layout(
        title="New vs Returning Customers per Month (full history baseline)",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        barmode="stack",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=month_order),
        yaxis=dict(title="Customers"),
        margin=dict(t=50, b=80, l=70, r=20), height=320
    )

    # ── 1b. Customer Repeat Frequency & Reorder Prediction ───────────────────
    last_data_date = df["Date"].max()

    inv_dates = (
        df.groupby(["Customer", "Doc No"])["Date"].min()
        .reset_index().sort_values(["Customer", "Date"])
    )
    cust_agent = df.groupby("Customer")["Agent"].agg(lambda x: x.mode()[0] if not x.mode().empty else "—")

    freq_rows = []
    for cust, grp in inv_dates.groupby("Customer"):
        dates = sorted(grp["Date"].tolist())
        n_orders = len(dates)
        last_order = dates[-1]
        if n_orders >= 2:
            gaps = [(dates[i+1] - dates[i]).days for i in range(len(dates)-1)]
            avg_gap = round(sum(gaps) / len(gaps))
            predicted_next = last_order + pd.Timedelta(days=avg_gap)
            days_until = (predicted_next - last_data_date).days
        else:
            avg_gap = None
            predicted_next = None
            days_until = None
        freq_rows.append({
            "Customer":      cust,
            "Agent":         cust_agent.get(cust, "—"),
            "Orders":        n_orders,
            "LastOrder":     last_order,
            "AvgGapDays":    avg_gap,
            "PredictedNext": predicted_next,
            "DaysUntil":     days_until,
        })

    freq_df = pd.DataFrame(freq_rows)
    freq_df_known = freq_df[
        freq_df["DaysUntil"].notna() &
        (freq_df["Orders"] >= 5) &
        (freq_df["AvgGapDays"] >= 10)
    ].copy()
    freq_due   = freq_df_known[freq_df_known["DaysUntil"] <= 30].sort_values("DaysUntil")
    freq_watch = freq_df_known[(freq_df_known["DaysUntil"] > 30) & (freq_df_known["DaysUntil"] <= 90)].sort_values("DaysUntil")

    def freq_row_html(r):
        days = int(r["DaysUntil"])
        if days < 0:
            badge_color, badge_text = "#e74c3c", f"Overdue {abs(days)}d"
        elif days == 0:
            badge_color, badge_text = "#e67e22", "Due Today"
        else:
            badge_color, badge_text = "#27ae60", f"In {days}d"
        agt = r.get("Agent", "—")
        agt_color = COLORS.get(agt, "#555")
        return (
            f'<td style="padding:6px 10px">{r["Customer"]}</td>'
            f'<td style="padding:6px 10px;text-align:center;font-weight:700;color:{agt_color}">{agt}</td>'
            f'<td style="padding:6px 10px;text-align:center">{int(r["Orders"])}</td>'
            f'<td style="padding:6px 10px;text-align:center">{r["LastOrder"].strftime("%d %b %Y")}</td>'
            f'<td style="padding:6px 10px;text-align:center">{int(r["AvgGapDays"])}d</td>'
            f'<td style="padding:6px 10px;text-align:center">{r["PredictedNext"].strftime("%d %b %Y")}</td>'
            f'<td style="padding:6px 10px;text-align:center">'
            f'<span style="background:{badge_color};color:white;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">{badge_text}</span>'
            f'</td>'
        )

    def freq_table_html(title, df_sub, header_color):
        if df_sub.empty:
            body = '<tr><td colspan="7" style="padding:12px;text-align:center;color:#aaa">—</td></tr>'
        else:
            body = "".join(
                f'<tr style="background:{"#f8f9fa" if i%2==0 else CARD_BG}">{freq_row_html(r)}</tr>'
                for i, (_, r) in enumerate(df_sub.iterrows(), 1)
            )
        return f"""
        <div style="margin-bottom:20px">
          <div style="font-size:13px;font-weight:700;color:{header_color};margin-bottom:6px">{title}</div>
          <table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead><tr style="background:{header_color};color:white">
              <th style="padding:6px 10px;text-align:left">Customer</th>
              <th style="padding:6px 10px">Agent</th>
              <th style="padding:6px 10px">Orders</th>
              <th style="padding:6px 10px">Last Order</th>
              <th style="padding:6px 10px">Avg Frequency</th>
              <th style="padding:6px 10px">Predicted Next</th>
              <th style="padding:6px 10px">Status</th>
            </tr></thead>
            <tbody>{body}</tbody>
          </table>
        </div>"""

    reorder_html = (
        freq_table_html(f"Due within 30 days (as of {last_data_date.strftime('%d %b %Y')}) — {len(freq_due)} customers", freq_due, "#e74c3c") +
        freq_table_html(f"Coming up in 31–90 days — {len(freq_watch)} customers", freq_watch, "#e67e22")
    )

    # ── 2. Product Mix by Top 6 Customers ─────────────────────────────────────
    top6_cust = df.groupby("Customer")["Revenue_RM"].sum().nlargest(6).index.tolist()
    cust_brand = df[df["Customer"].isin(top6_cust)].groupby(["Customer","Brand"])["Revenue_RM"].sum().reset_index()
    all_brands  = cust_brand.groupby("Brand")["Revenue_RM"].sum().nlargest(6).index.tolist()
    cust_brand["BrandGroup"] = cust_brand["Brand"].apply(lambda b: b if b in all_brands else "Other")
    cust_brand = cust_brand.groupby(["Customer","BrandGroup"])["Revenue_RM"].sum().reset_index()

    fig_mix = go.Figure()
    for b in all_brands + ["Other"]:
        bd = cust_brand[cust_brand["BrandGroup"] == b]
        fig_mix.add_trace(go.Bar(
            name=b,
            x=bd["Customer"].tolist(), y=bd["Revenue_RM"].tolist(),
            marker_color=COLORS.get(b, "#95a5a6"),
            hovertemplate="%{x}<br>" + b + ": RM %{y:,.0f}<extra></extra>",
        ))
    fig_mix.update_layout(
        title="Product Mix (Brand) — Top 6 Customers",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        barmode="stack",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-20),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=100, l=90, r=20), height=380
    )

    # ── 3. Area Analysis ──────────────────────────────────────────────────────
    PLACE_STATE = {
        "PENANG": "Penang", "BUTTERWORTH": "Penang",
        "J.BAHRU": "Johor", "JOHOR BAHRU": "Johor", "PARIT JAWA": "Johor",
        "SEGAMAT": "Johor", "BATU PAHAT": "Johor", "MUAR": "Johor",
        "TAIPING": "Perak", "PERAK": "Perak", "IPOH": "Perak",
        "BENTONG": "Pahang", "KUANTAN": "Pahang",
        "SHAH ALAM": "Selangor", "PUCHONG": "Selangor", "KLANG": "Selangor",
        "SUBANG": "Selangor", "PETALING JAYA": "Selangor",
        "SIBU": "Sarawak", "KUCHING": "Sarawak", "MIRI": "Sarawak",
        "KOTA KINABALU": "Sabah", "SANDAKAN": "Sabah",
        "SEREMBAN": "Negeri Sembilan",
        "MELAKA": "Melaka", "MALACCA": "Melaka",
        "ALOR SETAR": "Kedah", "SUNGAI PETANI": "Kedah",
        "KOTA BHARU": "Kelantan",
        "KUALA TERENGGANU": "Terengganu",
    }

    def _norm_area(val):
        if pd.isna(val):
            return "Unknown"
        s = str(val).strip()
        if s in ("----", "", "-"):
            return "Unknown"
        su = s.upper()
        if su in PLACE_STATE:
            return PLACE_STATE[su]
        if "SL" in su:
            return "Selangor"
        if "KL" in su:
            return "Kuala Lumpur"
        if s.replace("-","").replace(" ","").isdigit():
            return "Others"
        for place, state in PLACE_STATE.items():
            if place in su:
                return state
        return "Others"

    area_df = df[df["Area"].notna()].copy()
    area_df["Region"] = area_df["Area"].apply(_norm_area)
    area = area_df.groupby("Region").agg(
        Revenue=("Revenue_RM","sum"),
        Customers=("Customer","nunique"),
        Invoices=("Doc No","nunique"),
    ).reset_index()
    area = area[area["Revenue"] > 0].sort_values("Revenue", ascending=True)

    fig_area = go.Figure(go.Bar(
        x=area["Revenue"].tolist(), y=area["Region"].tolist(), orientation="h",
        marker_color="#1a3c5e",
        customdata=list(zip(area["Customers"], area["Invoices"])),
        hovertemplate="%{y}<br>Revenue: RM %{x:,.0f}<br>Customers: %{customdata[0]}<br>Invoices: %{customdata[1]}<extra></extra>",
    ))
    fig_area.update_layout(
        title=f"Total Revenue by Region ({period_start}–{period_end})",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=20, l=150, r=20), height=max(300, len(area) * 36 + 80)
    )

    # ── 4. Credit Term Analysis ───────────────────────────────────────────────
    df_c = df[df["CREDITTERM"].notna()].copy()
    df_c["CreditTermNorm"] = df_c["CREDITTERM"].apply(_norm_credit)
    df_c = df_c[df_c["CreditTermNorm"].notna()]

    credit = df_c.groupby("CreditTermNorm").agg(
        Revenue=("Revenue_RM","sum"),
        Customers=("Customer","nunique"),
        Invoices=("Doc No","nunique"),
    ).reset_index()
    credit = credit[credit["Revenue"] > 0].copy()

    credit["_sort"] = credit["CreditTermNorm"].apply(_days_key)
    credit = credit.sort_values("_sort", ascending=True)

    fig_credit = go.Figure(go.Bar(
        x=credit["Revenue"].tolist(), y=credit["CreditTermNorm"].tolist(), orientation="h",
        marker_color="#8e44ad",
        customdata=list(zip(credit["Customers"], credit["Invoices"])),
        hovertemplate="%{y}<br>Revenue: RM %{x:,.0f}<br>Customers: %{customdata[0]}<br>Invoices: %{customdata[1]}<extra></extra>",
    ))
    fig_credit.update_layout(
        title="Revenue by Credit Term (Normalised)",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        yaxis=dict(categoryorder="array", categoryarray=credit["CreditTermNorm"].tolist()),
        margin=dict(t=50, b=20, l=100, r=20), height=max(280, len(credit) * 36 + 80)
    )

    # ── 5. Seasonality (uses full history, not just 2025+) ───────────────────
    df_s = df_full.copy()
    df_s["Month"]     = df_s["Date"].dt.month
    df_s["MonthName"] = df_s["Date"].dt.strftime("%b")
    df_s["Year"]      = df_s["Date"].dt.year
    month_year  = df_s.groupby(["Year","Month","MonthName"])["Revenue_RM"].sum().reset_index()
    month_count = month_year.groupby("Month")["Year"].count()
    month_total = month_year.groupby(["Month","MonthName"])["Revenue_RM"].sum().reset_index()
    month_total["AvgRevenue"] = month_total.apply(
        lambda r: r["Revenue_RM"] / month_count.get(r["Month"], 1), axis=1
    )
    month_total = month_total.sort_values("Month")

    fig_season = go.Figure(go.Bar(
        x=month_total["MonthName"].tolist(),
        y=month_total["AvgRevenue"].tolist(),
        marker_color=[
            "#e74c3c" if v == month_total["AvgRevenue"].max()
            else "#2980b9" for v in month_total["AvgRevenue"]
        ],
        hovertemplate="%{x}<br>Avg Revenue: RM %{y:,.0f}<extra></extra>",
    ))
    fig_season.update_layout(
        title="Seasonality — Average Monthly Revenue by Calendar Month (All Years)",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(type="category"),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=40, l=90, r=20), height=320
    )

    # ── 6. Customer Flow — Company Total (New & Lost per Year) ──────────────
    cust_first_year = df_full.groupby("Customer")["Year"].min()
    all_years_full  = sorted(df_full["Year"].dropna().astype(int).unique())
    display_years   = sorted(df["Year"].dropna().astype(int).unique())

    flow_data = []
    for year in display_years:
        yr_df    = df[df["Year"] == year]
        new_mask = yr_df["Customer"].map(cust_first_year).astype("Int64") == year
        new_rev  = yr_df[new_mask].groupby("Customer")["Revenue_RM"].sum().sort_values(ascending=False)

        prev = year - 1
        if prev in all_years_full:
            prev_custs = set(df_full[df_full["Year"] == prev]["Customer"].unique())
            curr_custs = set(df_full[df_full["Year"] == year]["Customer"].unique())
            lost_set   = prev_custs - curr_custs
            lost_rev   = (
                df_full[(df_full["Year"] == prev) & df_full["Customer"].isin(lost_set)]
                .groupby("Customer")["Revenue_RM"].sum().sort_values(ascending=False)
            )
        else:
            lost_rev = pd.Series(dtype=float)

        flow_data.append({"year": year, "new": new_rev, "lost": lost_rev})

    fig_flow = go.Figure()
    fig_flow.add_trace(go.Bar(
        x=[str(d["year"]) for d in flow_data],
        y=[len(d["new"]) for d in flow_data],
        name="New Customers", marker_color="#27ae60",
        hovertemplate="%{x}<br>New customers: %{y}<extra></extra>",
    ))
    fig_flow.add_trace(go.Bar(
        x=[str(d["year"]) for d in flow_data],
        y=[-len(d["lost"]) for d in flow_data],
        name="Lost Customers", marker_color="#e74c3c",
        hovertemplate="%{x}<br>Lost customers: %{y}<extra></extra>",
    ))
    net_change = [len(d["new"]) - len(d["lost"]) for d in flow_data]
    fig_flow.add_trace(go.Scatter(
        x=[str(d["year"]) for d in flow_data], y=net_change,
        name="Net Change", mode="lines+markers+text",
        text=[f"+{n}" if n >= 0 else str(n) for n in net_change],
        textposition="top center",
        line=dict(color="#1a3c5e", width=2.5), marker=dict(size=8),
        hovertemplate="%{x}<br>Net: %{y}<extra></extra>",
    ))
    fig_flow.update_layout(
        title="Customer Acquisition & Churn — Company Total",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        barmode="relative",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        yaxis=dict(title="Customers", zeroline=True, zerolinecolor="#ccc"),
        xaxis=dict(type="category"),
        margin=dict(t=50, b=40, l=70, r=20), height=340
    )

    no2 = '<tr><td colspan="3" style="padding:8px;color:#aaa;text-align:center">—</td></tr>'
    flow_tables_html = ""
    for d in flow_data:
        year = d["year"]

        def _flow_rows(series, rev_color):
            rows = ""
            for i, (c, v) in enumerate(series.items(), 1):
                bg = "#f8f9fa" if i % 2 == 0 else CARD_BG
                agt = cust_agent.get(c, "—")
                agt_c = COLORS.get(agt, "#555")
                rows += (
                    f'<tr style="background:{bg}">'
                    f'<td style="padding:6px 10px">{c}</td>'
                    f'<td style="padding:6px 10px;text-align:center;font-weight:700;color:{agt_c}">{agt}</td>'
                    f'<td style="padding:6px 10px;text-align:right;color:{rev_color};font-weight:600">RM {v:,.0f}</td>'
                    f'</tr>'
                )
            return rows

        new_rows  = _flow_rows(d["new"],  "#27ae60")
        lost_rows = _flow_rows(d["lost"], "#e74c3c")

        flow_tables_html += f"""
        <div style="margin-bottom:24px">
          <div style="font-size:14px;font-weight:700;color:#1a3c5e;margin-bottom:12px;
                      padding-bottom:4px;border-bottom:1px solid #ddd">{year}</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
            <div>
              <div style="font-size:13px;font-weight:700;color:#27ae60;margin-bottom:6px">▲ New Customers ({len(d["new"])})</div>
              <table style="width:100%;border-collapse:collapse;font-size:12px">
                <thead><tr style="background:#27ae60;color:white">
                  <th style="padding:6px 10px;text-align:left">Customer</th>
                  <th style="padding:6px 10px">Agent</th>
                  <th style="padding:6px 10px;text-align:right">Revenue {year}</th>
                </tr></thead>
                <tbody>{new_rows if new_rows else no2}</tbody>
              </table>
            </div>
            <div>
              <div style="font-size:13px;font-weight:700;color:#e74c3c;margin-bottom:6px">▼ Lost from {year-1} ({len(d["lost"])})</div>
              <table style="width:100%;border-collapse:collapse;font-size:12px">
                <thead><tr style="background:#e74c3c;color:white">
                  <th style="padding:6px 10px;text-align:left">Customer</th>
                  <th style="padding:6px 10px">Agent</th>
                  <th style="padding:6px 10px;text-align:right">Last Revenue ({year-1})</th>
                </tr></thead>
                <tbody>{lost_rows if lost_rows else no2}</tbody>
              </table>
            </div>
          </div>
        </div>"""

    # ── 7. In-House Agent Customer Flow (AGENT1 & AGENT2) ────────────────────
    IN_HOUSE     = ["AGENT1", "AGENT2"]
    AGT_COLORS   = {"AGENT1": {"new": "#27ae60", "lost": "#e74c3c"},
                    "AGENT2": {"new": "#2980b9", "lost": "#e67e22"}}

    agent_flow = {}
    for agent in IN_HOUSE:
        agt_full       = df_full[df_full["Agent"].astype(str) == agent]
        agt_first_year = agt_full.groupby("Customer")["Year"].min()
        agt_years_all  = sorted(agt_full["Year"].dropna().astype(int).unique())
        records = []
        for year in display_years:
            yr_agt   = df[(df["Agent"] == agent) & (df["Year"] == year)]
            new_mask = yr_agt["Customer"].map(agt_first_year).astype("Int64") == year
            new_rev  = yr_agt[new_mask].groupby("Customer")["Revenue_RM"].sum().sort_values(ascending=False)

            prev = year - 1
            if prev in agt_years_all:
                prev_c   = set(agt_full[agt_full["Year"] == prev]["Customer"].unique())
                curr_c   = set(agt_full[agt_full["Year"] == year]["Customer"].unique())
                lost_rev = (
                    agt_full[(agt_full["Year"] == prev) & agt_full["Customer"].isin(prev_c - curr_c)]
                    .groupby("Customer")["Revenue_RM"].sum().sort_values(ascending=False)
                )
            else:
                lost_rev = pd.Series(dtype=float)

            records.append({"year": year, "new": new_rev, "lost": lost_rev})
        agent_flow[agent] = records

    fig_agent_flow = go.Figure()
    for agent in IN_HOUSE:
        recs = agent_flow[agent]
        fig_agent_flow.add_trace(go.Bar(
            x=[str(r["year"]) for r in recs],
            y=[len(r["new"]) for r in recs],
            name=f"{agent} New", marker_color=AGT_COLORS[agent]["new"],
            hovertemplate="%{x} — " + agent + " New: %{y}<extra></extra>",
        ))
        fig_agent_flow.add_trace(go.Bar(
            x=[str(r["year"]) for r in recs],
            y=[-len(r["lost"]) for r in recs],
            name=f"{agent} Lost", marker_color=AGT_COLORS[agent]["lost"],
            customdata=[len(r["lost"]) for r in recs],
            hovertemplate="%{x} — " + agent + " Lost: %{customdata}<extra></extra>",
        ))
    fig_agent_flow.update_layout(
        title="In-House Agent — Customer Acquisition & Churn by Year",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        yaxis=dict(title="Customers", zeroline=True, zerolinecolor="#ccc"),
        xaxis=dict(type="category"),
        margin=dict(t=50, b=40, l=70, r=20), height=360
    )

    no2a = '<tr><td colspan="2" style="padding:8px;color:#aaa;text-align:center">—</td></tr>'
    agent_flow_tables_html = ""
    for year in display_years:
        year_block = f"""
        <div style="margin-bottom:24px">
          <div style="font-size:14px;font-weight:700;color:#1a3c5e;margin-bottom:12px;
                      padding-bottom:4px;border-bottom:1px solid #ddd">{year}</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">"""
        for agent in IN_HOUSE:
            rec = next(r for r in agent_flow[agent] if r["year"] == year)
            ac  = AGT_COLORS[agent]
            new_rows = "".join(
                f'<tr style="background:{"#f8f9fa" if i%2==0 else CARD_BG}">'
                f'<td style="padding:5px 8px">{c}</td>'
                f'<td style="padding:5px 8px;text-align:right;color:{ac["new"]};font-weight:600">RM {v:,.0f}</td></tr>'
                for i, (c, v) in enumerate(rec["new"].items(), 1)
            )
            lost_rows = "".join(
                f'<tr style="background:{"#f8f9fa" if i%2==0 else CARD_BG}">'
                f'<td style="padding:5px 8px">{c}</td>'
                f'<td style="padding:5px 8px;text-align:right;color:{ac["lost"]};font-weight:600">RM {v:,.0f}</td></tr>'
                for i, (c, v) in enumerate(rec["lost"].items(), 1)
            )
            year_block += f"""
            <div>
              <div style="font-size:13px;font-weight:700;margin-bottom:8px;color:{ac["new"]}">{agent}</div>
              <div style="font-size:12px;font-weight:600;color:#27ae60;margin-bottom:4px">▲ New ({len(rec["new"])})</div>
              <table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:10px">
                <thead><tr style="background:{ac["new"]};color:white">
                  <th style="padding:5px 8px;text-align:left">Customer</th>
                  <th style="padding:5px 8px;text-align:right">Revenue {year}</th>
                </tr></thead>
                <tbody>{new_rows if new_rows else no2a}</tbody>
              </table>
              <div style="font-size:12px;font-weight:600;color:#e74c3c;margin-bottom:4px">▼ Lost from {year-1} ({len(rec["lost"])})</div>
              <table style="width:100%;border-collapse:collapse;font-size:11px">
                <thead><tr style="background:{ac["lost"]};color:white">
                  <th style="padding:5px 8px;text-align:left">Customer</th>
                  <th style="padding:5px 8px;text-align:right">Last Revenue ({year-1})</th>
                </tr></thead>
                <tbody>{lost_rows if lost_rows else no2a}</tbody>
              </table>
            </div>"""
        year_block += "</div></div>"
        agent_flow_tables_html += year_block

    # ── OEM vs Own Brand ──────────────────────────────────────────────────────

    bt = df.groupby("BrandType").agg(
        Revenue=("Revenue_RM","sum"),
        Margin=("Margin_RM","sum"),
        Customers=("Customer","nunique"),
        Invoices=("Doc No","nunique"),
    ).reset_index()
    bt["MarginPct"] = (bt["Margin"] / bt["Revenue"] * 100).round(1)
    bt["RevShare"]  = (bt["Revenue"] / bt["Revenue"].sum() * 100).round(1)

    def _bt(name):
        row = bt[bt["BrandType"] == name]
        return row.iloc[0] if len(row) else None

    oem_row = _bt("OEM")
    own_row = _bt("Own Brand")
    ind_row = _bt("Industrial Range")

    oem_cards = "".join([
        card_html("OEM Revenue",             fmt_rm(oem_row["Revenue"]) if oem_row is not None else "—",    f"{oem_row['RevShare']:.1f}% of total" if oem_row is not None else "", OEM_COLOR),
        card_html("OEM Margin %",            f"{oem_row['MarginPct']:.1f}%" if oem_row is not None else "—", f"{oem_row['Customers']:.0f} customers" if oem_row is not None else "", OEM_COLOR),
        card_html("Own Brand Revenue",       fmt_rm(own_row["Revenue"]) if own_row is not None else "—",    f"{own_row['RevShare']:.1f}% of total" if own_row is not None else "", OWN_COLOR),
        card_html("Own Brand Margin %",      f"{own_row['MarginPct']:.1f}%" if own_row is not None else "—", f"{own_row['Customers']:.0f} customers" if own_row is not None else "", OWN_COLOR),
        card_html("Industrial Range Revenue",fmt_rm(ind_row["Revenue"]) if ind_row is not None else "—",    f"{ind_row['RevShare']:.1f}% of total" if ind_row is not None else "", IND_COLOR),
        card_html("Industrial Margin %",     f"{ind_row['MarginPct']:.1f}%" if ind_row is not None else "—", f"{ind_row['Customers']:.0f} customers" if ind_row is not None else "", IND_COLOR),
    ])

    # Monthly trend OEM vs Own Brand
    bt_monthly = df.groupby(["YearMonth","MonthLabel","BrandType"]).agg(
        Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum")
    ).reset_index().sort_values("YearMonth")
    bt_monthly["MarginPct"] = (bt_monthly["Margin"] / bt_monthly["Revenue"] * 100).round(1)
    sorted_mo = bt_monthly.drop_duplicates("YearMonth").sort_values("YearMonth")["MonthLabel"].tolist()

    fig_bt_trend = go.Figure()
    for btype, color in [("Own Brand", OWN_COLOR), ("OEM", OEM_COLOR), ("Industrial Range", IND_COLOR)]:
        sub = bt_monthly[bt_monthly["BrandType"] == btype].set_index("MonthLabel")
        fig_bt_trend.add_trace(go.Bar(
            x=sorted_mo,
            y=[float(sub.loc[m,"Revenue"]) if m in sub.index else 0 for m in sorted_mo],
            name=btype, marker_color=color,
            hovertemplate="%{x}<br>" + btype + ": RM %{y:,.0f}<extra></extra>",
        ))
    fig_bt_trend.update_layout(
        title="Monthly Revenue — Own Brand vs OEM vs Industrial Range",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG, barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=sorted_mo),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=80, l=90, r=20), height=340
    )

    # Margin % trend by brand type
    fig_bt_mgn = go.Figure()
    for btype, color in [("Own Brand", OWN_COLOR), ("OEM", OEM_COLOR), ("Industrial Range", IND_COLOR)]:
        sub = bt_monthly[bt_monthly["BrandType"] == btype].set_index("MonthLabel")
        fig_bt_mgn.add_trace(go.Scatter(
            x=sorted_mo,
            y=[float(sub.loc[m,"MarginPct"]) if m in sub.index else None for m in sorted_mo],
            name=btype, mode="lines+markers",
            line=dict(color=color, width=2.5), marker=dict(size=6),
            hovertemplate="%{x}<br>" + btype + " margin: %{y:.1f}%%<extra></extra>",
        ))
    fig_bt_mgn.update_layout(
        title="Monthly Gross Margin % — Own Brand vs OEM vs Industrial Range",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=sorted_mo),
        yaxis=dict(ticksuffix="%", tickformat=".1f"),
        margin=dict(t=50, b=80, l=70, r=20), height=300
    )

    # Brand-level breakdown within each type
    brand_bt = df.groupby(["BrandType","Brand"]).agg(
        Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum")
    ).reset_index()
    brand_bt["MarginPct"] = (brand_bt["Margin"] / brand_bt["Revenue"] * 100).round(1)

    def _brand_table(btype, hcolor):
        sub = brand_bt[brand_bt["BrandType"] == btype].sort_values("Revenue", ascending=False)
        total_rev = sub["Revenue"].sum()
        rows = "".join(
            f'<tr style="background:{"#f8f9fa" if i%2==0 else CARD_BG}">'
            f'<td style="padding:7px 10px">{r["Brand"].replace(" (OEM)","")}</td>'
            f'<td style="padding:7px 10px;text-align:right;font-weight:600">RM {r["Revenue"]:,.0f}</td>'
            f'<td style="padding:7px 10px;text-align:right">{r["Revenue"]/total_rev*100:.1f}%</td>'
            f'<td style="padding:7px 10px;text-align:right;color:{"#27ae60" if r["MarginPct"]>=45 else "#e74c3c"};font-weight:600">{r["MarginPct"]:.1f}%</td>'
            f'</tr>'
            for i, (_, r) in enumerate(sub.iterrows(), 1)
        )
        return f"""
        <div>
          <div style="font-size:13px;font-weight:700;color:{hcolor};margin-bottom:8px">{btype}</div>
          <table style="width:100%;border-collapse:collapse;font-size:12px">
            <thead><tr style="background:{hcolor};color:white">
              <th style="padding:7px 10px;text-align:left">Brand</th>
              <th style="padding:7px 10px;text-align:right">Revenue</th>
              <th style="padding:7px 10px;text-align:right">% Share</th>
              <th style="padding:7px 10px;text-align:right">Margin %</th>
            </tr></thead>
            <tbody>{rows}</tbody>
          </table>
        </div>"""

    bt_brand_tables = f"""
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:16px">
      {_brand_table("Own Brand", OWN_COLOR)}
      {_brand_table("OEM", OEM_COLOR)}
    </div>
    <div style="display:grid;grid-template-columns:1fr;gap:20px">
      {_brand_table("Industrial Range", IND_COLOR)}
    </div>"""

    # ── A. Margin Trend ───────────────────────────────────────────────────────
    mgn_monthly = df.groupby(["YearMonth","MonthLabel"]).agg(
        Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum")
    ).reset_index().sort_values("YearMonth")
    mgn_monthly["MarginPct"] = (mgn_monthly["Margin"] / mgn_monthly["Revenue"] * 100).round(1)
    avg_mgn_pct = float(mgn_monthly["MarginPct"].mean())

    fig_mgn_trend = go.Figure()
    fig_mgn_trend.add_trace(go.Bar(
        x=mgn_monthly["MonthLabel"].tolist(), y=mgn_monthly["Margin"].tolist(),
        name="Gross Margin (RM)", marker_color="#27ae60", opacity=0.6,
        hovertemplate="%{x}<br>Margin: RM %{y:,.0f}<extra></extra>",
        yaxis="y1",
    ))
    fig_mgn_trend.add_trace(go.Scatter(
        x=mgn_monthly["MonthLabel"].tolist(), y=mgn_monthly["MarginPct"].tolist(),
        name="Margin %", mode="lines+markers+text",
        text=[f"{v:.1f}%" for v in mgn_monthly["MarginPct"]],
        textposition="top center", textfont=dict(size=10),
        line=dict(color="#1a3c5e", width=2.5), marker=dict(size=6),
        hovertemplate="%{x}<br>Margin %%: %{y:.1f}%%<extra></extra>",
        yaxis="y2",
    ))
    fig_mgn_trend.add_trace(go.Scatter(
        x=mgn_monthly["MonthLabel"].tolist(),
        y=[avg_mgn_pct] * len(mgn_monthly),
        name=f"Avg {avg_mgn_pct:.1f}%",
        mode="lines",
        line=dict(color="#e74c3c", dash="dash", width=1.5),
        hovertemplate=f"Average margin: {avg_mgn_pct:.1f}%<extra></extra>",
        yaxis="y2",
    ))
    fig_mgn_trend.update_layout(
        title="Monthly Gross Margin — RM & % Trend",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f", title="Margin RM"),
        yaxis2=dict(ticksuffix="%", title="Margin %", overlaying="y", side="right",
                    showgrid=False, range=[0, max(mgn_monthly["MarginPct"]) * 1.3]),
        margin=dict(t=50, b=80, l=90, r=70), height=360
    )

    # ── B. Revenue at Risk ────────────────────────────────────────────────────
    latest_date    = df["Date"].max()
    last_order     = df.groupby("Customer")["Date"].max()
    days_inactive  = (latest_date - last_order).dt.days
    twelve_mo_ago  = latest_date - pd.DateOffset(months=12)
    rev_12m        = df[df["Date"] >= twelve_mo_ago].groupby("Customer")["Revenue_RM"].sum()

    risk_agent = df.groupby("Customer")["Agent"].agg(lambda x: x.mode()[0] if not x.mode().empty else "—")

    risk_df = pd.DataFrame({
        "Customer":   days_inactive.index,
        "DaysSince":  days_inactive.values,
        "LastOrder":  last_order.values,
        "Revenue12m": [float(rev_12m.get(c, 0)) for c in days_inactive.index],
        "Agent":      [risk_agent.get(c, "—") for c in days_inactive.index],
    })
    risk_df = risk_df[risk_df["DaysSince"] >= 90].sort_values("DaysSince", ascending=False).copy()

    def _risk_level(d):
        if d >= 180: return "Critical (180d+)", "#e74c3c"
        return "Watch (90–179d)", "#e67e22"

    risk_total    = float(risk_df["Revenue12m"].sum())
    risk_critical = float(risk_df[risk_df["DaysSince"] >= 180]["Revenue12m"].sum())
    risk_watch    = float(risk_df[(risk_df["DaysSince"] >= 90) & (risk_df["DaysSince"] < 180)]["Revenue12m"].sum())

    risk_summary_cards = "".join([
        card_html("Customers Inactive 90d+", f"{len(risk_df)}", "need follow-up", "#e74c3c"),
        card_html("Revenue at Risk (12m)", fmt_rm(risk_total), "from inactive customers", "#e67e22"),
        card_html("Critical (180d+)", fmt_rm(risk_critical), f"{len(risk_df[risk_df['DaysSince']>=180])} customers", "#e74c3c"),
        card_html("Watch (90–179d)",  fmt_rm(risk_watch),    f"{len(risk_df[(risk_df['DaysSince']>=90)&(risk_df['DaysSince']<180)])} customers", "#e67e22"),
    ])

    no_risk = '<tr><td colspan="6" style="padding:10px;text-align:center;color:#aaa">No customers in this category</td></tr>'
    risk_rows = ""
    for i, (_, r) in enumerate(risk_df.iterrows(), 1):
        label, color = _risk_level(int(r["DaysSince"]))
        bg = "#f8f9fa" if i % 2 == 0 else CARD_BG
        agt_color = COLORS.get(r["Agent"], "#555")
        risk_rows += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:7px 10px">{r["Customer"]}</td>'
            f'<td style="padding:7px 10px;text-align:center;font-weight:700;color:{agt_color}">{r["Agent"]}</td>'
            f'<td style="padding:7px 10px;text-align:center">{r["LastOrder"].strftime("%d %b %Y")}</td>'
            f'<td style="padding:7px 10px;text-align:center">{int(r["DaysSince"])}d</td>'
            f'<td style="padding:7px 10px;text-align:right;font-weight:600">RM {r["Revenue12m"]:,.0f}</td>'
            f'<td style="padding:7px 10px;text-align:center">'
            f'<span style="background:{color};color:white;'
            f'padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">{label}</span>'
            f'</td></tr>'
        )
    risk_table_html = f"""
    <div class="tbl-wrap">
    <table style="width:100%;border-collapse:collapse;font-size:12px">
      <thead><tr style="background:#1a3c5e;color:white">
        <th style="padding:8px 10px;text-align:left">Customer</th>
        <th style="padding:8px 10px">Agent</th>
        <th style="padding:8px 10px">Last Order</th>
        <th style="padding:8px 10px">Days Inactive</th>
        <th style="padding:8px 10px;text-align:right">Revenue (12m)</th>
        <th style="padding:8px 10px">Status</th>
      </tr></thead>
      <tbody>{risk_rows if risk_rows else no_risk}</tbody>
    </table></div>"""

    # ── C. 80/20 Pareto Analysis ──────────────────────────────────────────────
    # Customer Pareto
    cp = df.groupby("Customer")["Revenue_RM"].sum().sort_values(ascending=False).reset_index()
    cp["CumPct"]  = cp["Revenue_RM"].cumsum() / cp["Revenue_RM"].sum() * 100
    cp["CustPct"] = (cp.index + 1) / len(cp) * 100
    cp80 = cp[cp["CumPct"] <= 80]
    n_cust_80 = len(cp80) + 1
    pct_cust_80 = round(n_cust_80 / len(cp) * 100, 1)

    fig_pareto_cust = go.Figure()
    fig_pareto_cust.add_trace(go.Scatter(
        x=cp["CustPct"].tolist(), y=cp["CumPct"].tolist(),
        mode="lines", name="Cumulative Revenue %",
        line=dict(color="#1a3c5e", width=2.5),
        hovertemplate="Top %{x:.1f}% customers → %{y:.1f}% revenue<extra></extra>",
    ))
    fig_pareto_cust.add_hline(y=80, line_dash="dash", line_color="#e74c3c",
        annotation_text="80% revenue", annotation_position="top left")
    fig_pareto_cust.add_vline(x=pct_cust_80, line_dash="dash", line_color="#e67e22",
        annotation_text=f"{pct_cust_80}% of customers", annotation_position="top right")
    fig_pareto_cust.update_layout(
        title=f"Customer Pareto — {pct_cust_80}% of customers = 80% of revenue ({n_cust_80} of {len(cp)})",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(ticksuffix="%", title="Cumulative % of Customers"),
        yaxis=dict(ticksuffix="%", title="Cumulative % of Revenue"),
        margin=dict(t=60, b=50, l=70, r=20), height=340
    )

    # Product Pareto — Own Brand vs OEM
    fig_pareto_prod = go.Figure()
    pareto_legend = []
    for btype, color in [("Own Brand", OWN_COLOR), ("OEM", OEM_COLOR)]:
        pp = (df[df["BrandType"] == btype]
              .groupby("Product")["Revenue_RM"].sum()
              .sort_values(ascending=False).reset_index())
        if pp.empty: continue
        pp["CumPct"]  = pp["Revenue_RM"].cumsum() / pp["Revenue_RM"].sum() * 100
        pp["ProdPct"] = (pp.index + 1) / len(pp) * 100
        n80 = len(pp[pp["CumPct"] <= 80]) + 1
        pct80 = round(n80 / len(pp) * 100, 1)
        pareto_legend.append(f"{btype}: top {pct80}% SKUs = 80% revenue ({n80}/{len(pp)})")
        fig_pareto_prod.add_trace(go.Scatter(
            x=pp["ProdPct"].tolist(), y=pp["CumPct"].tolist(),
            mode="lines", name=btype,
            line=dict(color=color, width=2.5),
            hovertemplate=f"{btype} — top %{{x:.1f}}% SKUs → %{{y:.1f}}% revenue<extra></extra>",
        ))
    fig_pareto_prod.add_hline(y=80, line_dash="dash", line_color="#e74c3c",
        annotation_text="80% revenue", annotation_position="top left")
    fig_pareto_prod.update_layout(
        title="Product Pareto — Own Brand vs OEM  |  " + "   ·   ".join(pareto_legend),
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(ticksuffix="%", title="Cumulative % of SKUs"),
        yaxis=dict(ticksuffix="%", title="Cumulative % of Revenue"),
        margin=dict(t=70, b=50, l=70, r=20), height=360
    )

    # ── F. Ecommerce vs Traditional ───────────────────────────────────────────
    ec = df.groupby("SalesChannel").agg(
        Revenue=("Revenue_RM","sum"),
        Margin=("Margin_RM","sum"),
        Customers=("Customer","nunique"),
        Invoices=("Doc No","nunique"),
    ).reset_index()
    ec["MarginPct"] = (ec["Margin"] / ec["Revenue"] * 100).round(1)
    ec["RevShare"]  = (ec["Revenue"] / ec["Revenue"].sum() * 100).round(1)

    def _ec(ch):
        row = ec[ec["SalesChannel"] == ch]
        return row.iloc[0] if len(row) else None

    ec_row   = _ec("Ecommerce")
    trad_row = _ec("Traditional")

    ecomm_cards = "".join([
        card_html("Ecommerce Revenue",    fmt_rm(ec_row["Revenue"]) if ec_row is not None else "—",
                  f"{ec_row['RevShare']:.1f}% of total" if ec_row is not None else "", ECOMM_COLOR),
        card_html("Ecommerce Margin %",   f"{ec_row['MarginPct']:.1f}%" if ec_row is not None else "—",
                  f"{int(ec_row['Customers'])} customers" if ec_row is not None else "", ECOMM_COLOR),
        card_html("Traditional Revenue",  fmt_rm(trad_row["Revenue"]) if trad_row is not None else "—",
                  f"{trad_row['RevShare']:.1f}% of total" if trad_row is not None else "", TRAD_COLOR),
        card_html("Traditional Margin %", f"{trad_row['MarginPct']:.1f}%" if trad_row is not None else "—",
                  f"{int(trad_row['Customers'])} customers" if trad_row is not None else "", TRAD_COLOR),
    ])

    # Monthly revenue trend: Ecommerce vs Traditional
    ec_monthly = df.groupby(["YearMonth","MonthLabel","SalesChannel"]).agg(
        Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum")
    ).reset_index().sort_values("YearMonth")
    ec_monthly["MarginPct"] = (ec_monthly["Margin"] / ec_monthly["Revenue"] * 100).round(1)
    ec_sorted_mo = ec_monthly.drop_duplicates("YearMonth").sort_values("YearMonth")["MonthLabel"].tolist()

    fig_ec_trend = go.Figure()
    for ch, color in [("Traditional", TRAD_COLOR), ("Ecommerce", ECOMM_COLOR)]:
        sub = ec_monthly[ec_monthly["SalesChannel"] == ch].set_index("MonthLabel")
        fig_ec_trend.add_trace(go.Bar(
            x=ec_sorted_mo,
            y=[float(sub.loc[m,"Revenue"]) if m in sub.index else 0 for m in ec_sorted_mo],
            name=ch, marker_color=color,
            hovertemplate="%{x}<br>" + ch + ": RM %{y:,.0f}<extra></extra>",
        ))
    fig_ec_trend.update_layout(
        title="Monthly Revenue — Ecommerce vs Traditional",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG, barmode="stack",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=ec_sorted_mo),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=80, l=90, r=20), height=340
    )

    # Ecommerce platform breakdown (Shopee / Lazada / TikTok)
    def _platform(customer):
        cu = str(customer).upper()
        if "SHOPEE"  in cu: return "Shopee"
        if "LAZADA"  in cu: return "Lazada"
        if "TIKTOK"  in cu: return "TikTok"
        return "Other Ecommerce"

    ec_df = df[df["SalesChannel"] == "Ecommerce"].copy()
    ec_df["Platform"] = ec_df["Customer"].apply(_platform)

    plat_sum = ec_df.groupby("Platform").agg(
        Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum"),
        Customers=("Customer","nunique"), Invoices=("Doc No","nunique"),
    ).reset_index().sort_values("Revenue", ascending=False)
    plat_sum["MarginPct"] = (plat_sum["Margin"] / plat_sum["Revenue"] * 100).round(1)
    plat_sum["Share"]     = (plat_sum["Revenue"] / plat_sum["Revenue"].sum() * 100).round(1)

    PLAT_COLORS = {"Shopee":"#ee4d2d","Lazada":"#0f146b","TikTok":"#010101","Other Ecommerce":"#78909c"}

    fig_ec_plat = go.Figure()
    fig_ec_plat.add_trace(go.Bar(
        x=plat_sum["Platform"].tolist(), y=plat_sum["Revenue"].tolist(),
        marker_color=[PLAT_COLORS.get(p,"#bdc3c7") for p in plat_sum["Platform"]],
        text=[f"RM {v:,.0f}<br>{s:.1f}%" for v, s in zip(plat_sum["Revenue"], plat_sum["Share"])],
        textposition="outside",
        hovertemplate="%{x}<br>Revenue: RM %{y:,.0f}<extra></extra>",
    ))
    fig_ec_plat.update_layout(
        title="Ecommerce Revenue by Platform",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(type="category"),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        showlegend=False,
        margin=dict(t=50, b=40, l=90, r=20), height=320
    )

    # Ecommerce platform detail table
    ec_plat_rows = ""
    ec_total_rev = float(plat_sum["Revenue"].sum())
    for i, (_, r) in enumerate(plat_sum.iterrows(), 1):
        bg = "#f8f9fa" if i % 2 == 0 else CARD_BG
        mgn_color = "#27ae60" if r["MarginPct"] >= 45 else "#e74c3c"
        ec_plat_rows += (
            f'<tr style="background:{bg}">'
            f'<td style="padding:8px 12px;font-weight:700;color:{PLAT_COLORS.get(r["Platform"],"#333")}">{r["Platform"]}</td>'
            f'<td style="padding:8px 12px;text-align:right;font-weight:600">RM {r["Revenue"]:,.0f}</td>'
            f'<td style="padding:8px 12px;text-align:right">{r["Share"]:.1f}%</td>'
            f'<td style="padding:8px 12px;text-align:right;color:{mgn_color};font-weight:600">{r["MarginPct"]:.1f}%</td>'
            f'<td style="padding:8px 12px;text-align:right">{int(r["Customers"])}</td>'
            f'<td style="padding:8px 12px;text-align:right">{int(r["Invoices"])}</td>'
            f'</tr>'
        )
    ec_plat_html = f"""
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead><tr style="background:#1a3c5e;color:white">
        <th style="padding:8px 12px;text-align:left">Platform</th>
        <th style="padding:8px 12px;text-align:right">Revenue</th>
        <th style="padding:8px 12px;text-align:right">% Share</th>
        <th style="padding:8px 12px;text-align:right">Margin %</th>
        <th style="padding:8px 12px;text-align:right">Customers</th>
        <th style="padding:8px 12px;text-align:right">Invoices</th>
      </tr></thead>
      <tbody>{ec_plat_rows if ec_plat_rows else "<tr><td colspan='6' style='padding:12px;text-align:center;color:#aaa'>No ecommerce data</td></tr>"}</tbody>
    </table>"""

    # ── D. Product Performance Matrix ────────────────────────────────────────
    brand_perf = df.groupby("Brand").agg(
        Revenue=("Revenue_RM","sum"),
        Margin=("Margin_RM","sum"),
        Customers=("Customer","nunique"),
        Invoices=("Doc No","nunique"),
    ).reset_index()
    brand_perf = brand_perf[brand_perf["Revenue"] > 0].copy()
    brand_perf["MarginPct"] = (brand_perf["Margin"] / brand_perf["Revenue"] * 100).round(1)

    med_rev = float(brand_perf["Revenue"].median())
    med_mgn = float(brand_perf["MarginPct"].median())

    QUAD_META = {
        "⭐ Star":          {"color":"#27ae60", "bg":"rgba(39,174,96,0.07)",  "desc":"High Revenue · High Margin — Protect & Grow"},
        "💰 Cash Cow":      {"color":"#e67e22", "bg":"rgba(230,126,34,0.07)", "desc":"High Revenue · Low Margin — Optimise Cost"},
        "🚀 Question Mark": {"color":"#2980b9", "bg":"rgba(41,128,185,0.07)", "desc":"Low Revenue · High Margin — Invest & Push"},
        "⚠️ Dog":           {"color":"#e74c3c", "bg":"rgba(231,76,60,0.07)",  "desc":"Low Revenue · Low Margin — Review / Exit"},
    }

    def _quadrant(rev, mgn):
        if rev >= med_rev and mgn >= med_mgn: return "⭐ Star"
        if rev >= med_rev and mgn <  med_mgn: return "💰 Cash Cow"
        if rev <  med_rev and mgn >= med_mgn: return "🚀 Question Mark"
        return "⚠️ Dog"

    brand_perf["Quadrant"] = brand_perf.apply(lambda r: _quadrant(r["Revenue"], r["MarginPct"]), axis=1)
    brand_perf = brand_perf.sort_values("Revenue", ascending=False)

    # Scatter: log X so small brands are readable; label only top-5 per quadrant
    top5 = set(brand_perf.head(5)["Brand"])

    fig_matrix = go.Figure()
    for q, meta in QUAD_META.items():
        sub = brand_perf[brand_perf["Quadrant"] == q]
        if sub.empty: continue
        labels = [b if b in top5 else "" for b in sub["Brand"]]
        fig_matrix.add_trace(go.Scatter(
            x=sub["Revenue"].tolist(), y=sub["MarginPct"].tolist(),
            mode="markers+text", name=q,
            text=sub["Brand"].tolist(),
            texttemplate=[b if b in top5 else "" for b in sub["Brand"]],
            textposition="top center", textfont=dict(size=10),
            marker=dict(
                size=[max(16, min(44, int(c) * 3)) for c in sub["Customers"]],
                color=meta["color"], opacity=0.8, line=dict(width=1.5, color="white")
            ),
            customdata=sub[["Customers","Invoices","MarginPct","Revenue"]].values,
            hovertemplate=(
                "<b>%{customdata[0]:s}</b><br>"
                "Revenue: RM %{x:,.0f}<br>"
                "Margin: %{y:.1f}%<br>"
                "Customers: %{customdata[0]}<br>"
                "Invoices: %{customdata[1]}<extra>" + q + "</extra>"
            ),
        ))

    # Use the customdata trick for brand name in hover — fix via simpler approach
    fig_matrix = go.Figure()
    for q, meta in QUAD_META.items():
        sub = brand_perf[brand_perf["Quadrant"] == q].copy()
        if sub.empty: continue
        fig_matrix.add_trace(go.Scatter(
            x=sub["Revenue"].tolist(), y=sub["MarginPct"].tolist(),
            mode="markers+text", name=q,
            text=sub["Brand"].tolist(),
            texttemplate=[b if b in top5 else "" for b in sub["Brand"]],
            textposition="top center", textfont=dict(size=10, color="#333"),
            marker=dict(
                size=[max(16, min(44, int(c) * 3)) for c in sub["Customers"]],
                color=meta["color"], opacity=0.8, line=dict(width=1.5, color="white")
            ),
            customdata=list(zip(sub["Brand"], sub["Customers"], sub["Invoices"])),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Revenue: RM %{x:,.0f}<br>"
                "Margin: %{y:.1f}%<br>"
                "Customers: %{customdata[1]}<br>"
                "Invoices: %{customdata[2]}<extra>" + q + "</extra>"
            ),
        ))

    fig_matrix.add_vline(x=med_rev, line_dash="dot", line_color="#999", line_width=1.5,
        annotation_text=f"Median Rev: {fmt_rm(med_rev)}", annotation_position="top right",
        annotation_font_size=10, annotation_font_color="#666")
    fig_matrix.add_hline(y=med_mgn, line_dash="dot", line_color="#999", line_width=1.5,
        annotation_text=f"Median Margin: {med_mgn:.1f}%", annotation_position="top left",
        annotation_font_size=10, annotation_font_color="#666")
    fig_matrix.update_layout(
        title="Brand Performance Matrix — Revenue vs Margin %  (bubble size = no. of customers, hover for details)",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(type="log", tickprefix="RM ", tickformat=",.0f", title="Total Revenue (log scale)"),
        yaxis=dict(ticksuffix="%", title="Gross Margin %"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(t=65, b=60, l=80, r=20), height=500
    )

    # Quadrant summary table
    quad_rows = ""
    for q, meta in QUAD_META.items():
        sub = brand_perf[brand_perf["Quadrant"] == q].sort_values("Revenue", ascending=False)
        if sub.empty: continue
        brands_str = " · ".join(sub["Brand"].str.replace(" (OEM)","",regex=False).tolist())
        quad_rows += (
            f'<tr>'
            f'<td style="padding:10px 14px;font-weight:700;color:{meta["color"]};white-space:nowrap">{q}</td>'
            f'<td style="padding:10px 14px;color:#555;font-size:12px">{meta["desc"]}</td>'
            f'<td style="padding:10px 14px;font-size:12px">{brands_str}</td>'
            f'<td style="padding:10px 14px;text-align:right;font-weight:600">{len(sub)}</td>'
            f'</tr>'
        )
    matrix_summary_html = f"""
    <table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:16px">
      <thead><tr style="background:#1a3c5e;color:white">
        <th style="padding:10px 14px;text-align:left">Quadrant</th>
        <th style="padding:10px 14px;text-align:left">Strategy</th>
        <th style="padding:10px 14px;text-align:left">Brands</th>
        <th style="padding:10px 14px;text-align:right">Count</th>
      </tr></thead>
      <tbody>{quad_rows}</tbody>
    </table>"""

    # ── E. Average Order Value (AOV) Trend ────────────────────────────────────
    aov_monthly = df.groupby(["YearMonth","MonthLabel"]).agg(
        Revenue=("Revenue_RM","sum"),
        Invoices=("Doc No","nunique"),
    ).reset_index().sort_values("YearMonth")
    aov_monthly["AOV"] = aov_monthly["Revenue"] / aov_monthly["Invoices"]
    avg_aov = float(aov_monthly["AOV"].mean())
    aov_labels = aov_monthly["MonthLabel"].tolist()

    fig_aov = go.Figure()
    fig_aov.add_trace(go.Bar(
        x=aov_labels, y=aov_monthly["Invoices"].tolist(),
        name="No. of Invoices", marker_color="#bdc3c7", opacity=0.45,
        hovertemplate="%{x}<br>Invoices: %{y}<extra></extra>",
        yaxis="y2",
    ))
    fig_aov.add_trace(go.Scatter(
        x=aov_labels, y=aov_monthly["AOV"].tolist(),
        name="Avg Order Value (RM)", mode="lines+markers+text",
        text=[f"RM {v:,.0f}" for v in aov_monthly["AOV"]],
        textposition="top center", textfont=dict(size=9),
        line=dict(color="#1a3c5e", width=2.5), marker=dict(size=6),
        hovertemplate="%{x}<br>AOV: RM %{y:,.0f}<extra></extra>",
        yaxis="y1",
    ))
    fig_aov.add_trace(go.Scatter(
        x=aov_labels, y=[avg_aov] * len(aov_labels),
        name=f"Overall Avg: RM {avg_aov:,.0f}",
        mode="lines", line=dict(color="#e74c3c", dash="dash", width=1.5),
        hovertemplate=f"Overall avg AOV: RM {avg_aov:,.0f}<extra></extra>",
        yaxis="y1",
    ))
    fig_aov.update_layout(
        title="Average Order Value (AOV) per Month",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f", title="AOV (RM)"),
        yaxis2=dict(title="No. of Invoices", overlaying="y", side="right", showgrid=False),
        margin=dict(t=50, b=80, l=90, r=70), height=360
    )

    # ── G. RFM Customer Segmentation ─────────────────────────────────────────
    latest_rfm = df["Date"].max()
    rfm = df.groupby("Customer").agg(
        LastOrder =("Date",       "max"),
        Frequency =("Doc No",     "nunique"),
        Monetary  =("Revenue_RM", "sum"),
        Agent     =("Agent",      lambda x: x.mode()[0] if not x.mode().empty else "—"),
    ).reset_index()
    rfm["Recency"] = (latest_rfm - rfm["LastOrder"]).dt.days

    # Score 1–5 (5 = best). Use rank to break ties safely.
    rfm["R"] = pd.qcut(rfm["Recency"].rank(method="first"),  5, labels=[5,4,3,2,1]).astype(int)
    rfm["F"] = pd.qcut(rfm["Frequency"].rank(method="first"),5, labels=[1,2,3,4,5]).astype(int)
    rfm["M"] = pd.qcut(rfm["Monetary"].rank(method="first"), 5, labels=[1,2,3,4,5]).astype(int)

    SEG_META = {
        "🏆 Champion":    {"color":"#27ae60", "bg":"#eafaf1", "action":"VIP 维护，优先推介新品"},
        "💚 Loyal":       {"color":"#2980b9", "bg":"#eaf4fb", "action":"定期回访，交叉销售更多产品线"},
        "🌱 Promising":   {"color":"#16a085", "bg":"#e8f8f5", "action":"新客培养，扩大产品覆盖"},
        "💛 Potential":   {"color":"#d4ac0d", "bg":"#fef9e7", "action":"提升购买频率，增加下单次数"},
        "⚠️ At Risk":     {"color":"#e67e22", "bg":"#fef5ec", "action":"紧急跟进，了解沉寂原因"},
        "😴 Hibernating": {"color":"#7f8c8d", "bg":"#f2f3f4", "action":"低优先级，季度性激活"},
        "❌ Lost":        {"color":"#e74c3c", "bg":"#fdedec", "action":"考虑重新激活，或接受流失"},
    }

    def _segment(r, f, m):
        if r >= 5 and f >= 4 and m >= 4: return "🏆 Champion"
        if r >= 4 and f >= 3 and m >= 3: return "💚 Loyal"
        if r >= 4 and f <= 2:            return "🌱 Promising"
        if r >= 3 and f >= 2 and m >= 2: return "💛 Potential"
        if r <= 2 and f >= 3:            return "⚠️ At Risk"
        if r <= 2 and f <= 2 and m <= 2: return "😴 Hibernating"
        return "❌ Lost"

    rfm["Segment"] = rfm.apply(lambda row: _segment(row["R"], row["F"], row["M"]), axis=1)

    # ── Segment summary ───────────────────────────────────────────────────────
    seg_sum = rfm.groupby("Segment").agg(
        Customers=("Customer","count"),
        Revenue  =("Monetary","sum"),
        AvgR     =("Recency","mean"),
        AvgF     =("Frequency","mean"),
        AvgM     =("Monetary","mean"),
    ).reset_index()
    seg_sum["RevShare"] = (seg_sum["Revenue"] / seg_sum["Revenue"].sum() * 100).round(1)

    seg_order = list(SEG_META.keys())
    seg_sum["_ord"] = seg_sum["Segment"].map({s:i for i,s in enumerate(seg_order)})
    seg_sum = seg_sum.sort_values("_ord").drop(columns="_ord")

    # KPI cards: one per segment
    rfm_cards = ""
    for _, s in seg_sum.iterrows():
        meta = SEG_META.get(s["Segment"], {"color":"#333","bg":"#fff","action":""})
        rfm_cards += f"""
        <div style="background:{meta['bg']};border-left:4px solid {meta['color']};
                    border-radius:8px;padding:14px 16px;flex:1;min-width:150px">
          <div style="font-size:13px;font-weight:700;color:{meta['color']};margin-bottom:6px">{s['Segment']}</div>
          <div style="font-size:22px;font-weight:700;color:#1a3c5e">{int(s['Customers'])}</div>
          <div style="font-size:12px;color:#888;margin-top:2px">customers</div>
          <div style="font-size:12px;color:#555;margin-top:4px;font-weight:600">
            RM {s['Revenue']:,.0f} &nbsp;·&nbsp; {s['RevShare']:.1f}%</div>
        </div>"""

    # ── Scatter: R-score vs F-score, bubble = M ───────────────────────────────
    fig_rfm = go.Figure()
    for seg, meta in SEG_META.items():
        sub = rfm[rfm["Segment"] == seg]
        if sub.empty: continue
        fig_rfm.add_trace(go.Scatter(
            x=sub["R"].tolist(), y=sub["F"].tolist(),
            mode="markers", name=seg,
            marker=dict(
                size=[max(10, min(36, int(m/5000))) for m in sub["Monetary"]],
                color=meta["color"], opacity=0.75, line=dict(width=1, color="white")
            ),
            customdata=list(zip(
                sub["Customer"], sub["Recency"], sub["Frequency"],
                sub["Monetary"], sub["Agent"]
            )),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Recency: %{customdata[1]} days<br>"
                "Frequency: %{customdata[2]} invoices<br>"
                "Revenue: RM %{customdata[3]:,.0f}<br>"
                "Agent: %{customdata[4]}<extra>" + seg + "</extra>"
            ),
        ))
    fig_rfm.update_layout(
        title="RFM Segmentation — Recency Score vs Frequency Score  (bubble size = Revenue)",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
        xaxis=dict(title="Recency Score (5=most recent)", tickvals=[1,2,3,4,5],
                   range=[0.5,5.5], showgrid=True, gridcolor="#eee"),
        yaxis=dict(title="Frequency Score (5=most frequent)", tickvals=[1,2,3,4,5],
                   range=[0.5,5.5], showgrid=True, gridcolor="#eee"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(t=65, b=60, l=80, r=20), height=460
    )

    # ── Segment detail tables ─────────────────────────────────────────────────
    rfm_detail_html = ""
    for seg in seg_order:
        sub = rfm[rfm["Segment"] == seg].sort_values("Monetary", ascending=False)
        if sub.empty: continue
        meta = SEG_META[seg]
        rows = ""
        for i, (_, r) in enumerate(sub.iterrows(), 1):
            bg = "#f8f9fa" if i % 2 == 0 else CARD_BG
            rows += (
                f'<tr style="background:{bg}">'
                f'<td style="padding:6px 10px">{r["Customer"]}</td>'
                f'<td style="padding:6px 10px;text-align:center">{int(r["Recency"])}d</td>'
                f'<td style="padding:6px 10px;text-align:center">{int(r["Frequency"])}</td>'
                f'<td style="padding:6px 10px;text-align:right;font-weight:600">RM {r["Monetary"]:,.0f}</td>'
                f'<td style="padding:6px 10px;text-align:center">{r["R"]}-{r["F"]}-{r["M"]}</td>'
                f'<td style="padding:6px 10px;text-align:center;color:#555">{r["Agent"]}</td>'
                f'</tr>'
            )
        count = len(sub)
        total_rev = sub["Monetary"].sum()
        rfm_detail_html += f"""
        <details style="margin-bottom:12px">
          <summary style="cursor:pointer;padding:10px 14px;background:{meta['bg']};
                          border-left:4px solid {meta['color']};border-radius:6px;
                          font-weight:700;color:{meta['color']};list-style:none;
                          display:flex;justify-content:space-between;align-items:center">
            <span>{seg} &nbsp;·&nbsp; {count} customers &nbsp;·&nbsp; RM {total_rev:,.0f}</span>
            <span style="font-size:11px;color:#888;font-weight:400">{meta['action']}</span>
          </summary>
          <div style="padding:0 4px">
            <table style="width:100%;border-collapse:collapse;font-size:12px">
              <thead><tr style="background:{meta['color']};color:white">
                <th style="padding:7px 10px;text-align:left">Customer</th>
                <th style="padding:7px 10px">Days Since Order</th>
                <th style="padding:7px 10px">Invoices</th>
                <th style="padding:7px 10px;text-align:right">Total Revenue</th>
                <th style="padding:7px 10px">R-F-M</th>
                <th style="padding:7px 10px">Agent</th>
              </tr></thead>
              <tbody>{rows}</tbody>
            </table>
          </div>
        </details>"""

    # ── Assemble HTML ─────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>LLG Industries — Sales Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', sans-serif; background: {BG}; color: #333; }}
  .header {{ background: #1a3c5e; color: white; padding: 24px 32px; }}
  .header h1 {{ font-size: 22px; font-weight: 700; }}
  .header p  {{ font-size: 13px; opacity: .7; margin-top: 4px; }}
  .container {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}
  .cards  {{ display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }}
  .row    {{ display: grid; gap: 16px; margin-bottom: 16px; }}
  .row-2  {{ grid-template-columns: 1fr 1fr; }}
  .row-1  {{ grid-template-columns: 1fr; }}
  .panel  {{ background: {CARD_BG}; border-radius: 12px; padding: 8px;
             box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
  .panel-p {{ background: {CARD_BG}; border-radius: 12px; padding: 16px;
              box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
  .tbl-wrap {{ overflow-x: auto; -webkit-overflow-scrolling: touch; }}
  @media (max-width: 768px) {{
    .header {{ padding: 16px; }}
    .header h1 {{ font-size: 17px; }}
    .container {{ padding: 8px; overflow-x: hidden; }}
    .row-2 {{ grid-template-columns: 1fr; }}
    .panel, .panel-p {{ padding: 6px; }}
    .cards > div {{ min-width: calc(50% - 6px) !important; padding: 14px 10px !important; }}
    .cards > div > div:nth-child(2) {{ font-size: 20px !important; }}
    table {{ font-size: 11px; display: block; overflow-x: auto; -webkit-overflow-scrolling: touch; max-width: 100%; }}
    details summary {{ font-size: 13px; }}
  }}
  @media (max-width: 480px) {{
    .cards > div {{ min-width: 100% !important; }}
    .cards > div > div:nth-child(2) {{ font-size: 18px !important; }}
  }}
</style>
</head>
<body>
<div class="header">
  <h1>{report_title}</h1>
  <p>LLE + LLG &nbsp;|&nbsp; {period_start} – {period_end} &nbsp;|&nbsp;
     2025 onwards &nbsp;|&nbsp; Generated {datetime.now().strftime('%d %b %Y %H:%M')}</p>
</div>
<div class="container">

  {section_header("Overview")}
  <div class="cards">{cards}</div>

  {section_header(f"YTD Performance — {ytd_year} vs {ytd_year-1} (Jan – {MONTHS_SHORT[ytd_month-1]})")}
  <div class="cards">{ytd_cards}</div>
  <div class="row row-1"><div class="panel">{fig_html(fig_ytd)}</div></div>

  {section_header("Sales Trend")}
  <div class="row row-1"><div class="panel">{fig_html(fig_trend)}</div></div>
  <div class="row row-1"><div class="panel">{fig_html(fig_yoy)}</div></div>

  {section_header("Margin Trend")}
  <div class="row row-1"><div class="panel">{fig_html(fig_mgn_trend)}</div></div>

  {section_header("Seasonality")}
  <div class="row row-1"><div class="panel">{fig_html(fig_season)}</div></div>

  {section_header("Brand Type Analysis — Own Brand / OEM / Industrial Range")}
  <div class="cards">{oem_cards}</div>
  <div class="row row-2">
    <div class="panel">{fig_html(fig_bt_trend)}</div>
    <div class="panel">{fig_html(fig_bt_mgn)}</div>
  </div>
  <div class="row row-1"><div class="panel-p">{bt_brand_tables}</div></div>

  {section_header("Average Order Value (AOV)")}
  <div class="row row-1"><div class="panel">{fig_html(fig_aov)}</div></div>

  {section_header("Brand & Channel")}
  <div class="row row-2">
    <div class="panel">{fig_html(fig_brand_rev)}</div>
    <div class="panel">{fig_html(fig_brand_mgn)}</div>
  </div>
  <div class="row row-1"><div class="panel">{fig_html(fig_channel)}</div></div>

  {section_header("Brand Performance Matrix")}
  <div class="row row-1"><div class="panel">{fig_html(fig_matrix)}</div></div>
  <div class="row row-1"><div class="panel-p">{matrix_summary_html}</div></div>

  {section_header("Ecommerce vs Traditional")}
  <div class="cards">{ecomm_cards}</div>
  <div class="row row-2">
    <div class="panel">{fig_html(fig_ec_trend)}</div>
    <div class="panel">{fig_html(fig_ec_plat)}</div>
  </div>
  <div class="row row-1"><div class="panel-p">{ec_plat_html}</div></div>

  {section_header("In-House Sales Agent Analysis")}
  <div class="row row-1"><div class="panel-p">{agent_table}</div></div>
  <div class="row row-2">
    <div class="panel">{fig_html(fig_agent)}</div>
    <div class="panel">{fig_html(fig_agent_cust)}</div>
  </div>

  {section_header("Customer Analysis")}
  <div class="row row-2">
    <div class="panel">{fig_html(fig_repeat)}</div>
    <div class="panel">{fig_html(fig_mix)}</div>
  </div>
  <div class="row row-1"><div class="panel">{fig_html(fig_cust)}</div></div>

  {section_header("RFM Customer Segmentation")}
  <div style="display:flex;flex-wrap:wrap;gap:12px;padding:4px 0 16px">{rfm_cards}</div>
  <div class="row row-1"><div class="panel">{fig_html(fig_rfm)}</div></div>
  <div class="row row-1"><div class="panel-p">{rfm_detail_html}</div></div>

  {section_header("Reorder Prediction — Customers Due Soon")}
  <div class="row row-1"><div class="panel-p">{reorder_html}</div></div>

  {section_header("Customer Acquisition & Churn — Company Total")}
  <div class="row row-1"><div class="panel">{fig_html(fig_flow)}</div></div>
  <div class="row row-1"><div class="panel-p">{flow_tables_html}</div></div>

  {section_header("In-House Agent — Customer Acquisition & Churn")}
  <div class="row row-1"><div class="panel">{fig_html(fig_agent_flow)}</div></div>
  <div class="row row-1"><div class="panel-p">{agent_flow_tables_html}</div></div>

  {section_header("Geographic & Credit Term Analysis")}
  <div class="row row-2">
    <div class="panel">{fig_html(fig_area)}</div>
    <div class="panel">{fig_html(fig_credit)}</div>
  </div>

  {section_header("Top Products")}
  <div class="row row-1"><div class="panel-p">{prod_table}</div></div>

  {section_header("80/20 Pareto Analysis")}
  <div class="row row-2">
    <div class="panel">{fig_html(fig_pareto_cust)}</div>
    <div class="panel">{fig_html(fig_pareto_prod)}</div>
  </div>

  {section_header("Revenue at Risk — Inactive Customers")}
  <div class="cards">{risk_summary_cards}</div>
  <div class="row row-1"><div class="panel-p">{risk_table_html}</div></div>

</div>
</body>
</html>"""
    return html

# ── Agent Dashboard ────────────────────────────────────────────────────────────

def build_agent_dashboard(df: pd.DataFrame, foc_cost: float, df_full: pd.DataFrame,
                          agents: list = None) -> str:
    IN_HOUSE     = agents or ["AGENT1", "AGENT2"]
    AGENT_COLOR  = {"AGENT1": "#27ae60", "AGENT2": "#2980b9", "AGENT3": "#8e44ad"}
    LOST_COLOR   = {"AGENT1": "#e74c3c", "AGENT2": "#e67e22", "AGENT3": "#c0392b"}

    rev         = df["Revenue_RM"].sum()
    mgn         = df["Margin_RM"].sum()
    mgp         = mgn / rev * 100 if rev else 0
    n_months    = df["YearMonth"].nunique()
    period_start = df["Date"].min().strftime("%b %Y")
    period_end   = df["Date"].max().strftime("%b %Y")
    generated   = datetime.now().strftime("%d %b %Y %H:%M")

    # ── Overview Cards ────────────────────────────────────────────────────────
    overview_cards = "".join([
        card_html("Total Revenue",       fmt_rm(rev),            f"{period_start} – {period_end}", "#1a3c5e"),
        card_html("Total Gross Margin",  fmt_rm(mgn),            f"{mgp:.1f}% margin",             "#27ae60"),
        card_html("Avg Monthly Revenue", fmt_rm(rev / n_months), "per month",                      "#2980b9"),
        card_html("Invoices",            f"{df['Doc No'].nunique():,}", "unique documents",          "#8e44ad"),
        card_html("FOC Cost",            fmt_rm(foc_cost),       "free goods at cost",              "#e67e22"),
    ])

    # ── Overview: AGENT1 vs AGENT2 Monthly Comparison ────────────────────────
    monthly = df.groupby(["YearMonth","MonthLabel","Agent"]).agg(
        Revenue=("Revenue_RM","sum")
    ).reset_index().sort_values("YearMonth")
    sorted_months = monthly.drop_duplicates("YearMonth").sort_values("YearMonth")["MonthLabel"].tolist()

    fig_ov_monthly = go.Figure()
    for agent in IN_HOUSE:
        m = monthly[monthly["Agent"] == agent].set_index("MonthLabel")
        fig_ov_monthly.add_trace(go.Bar(
            x=sorted_months,
            y=[float(m.loc[ml, "Revenue"]) if ml in m.index else 0 for ml in sorted_months],
            name=agent, marker_color=AGENT_COLOR[agent],
            hovertemplate="%{x}<br>" + agent + ": RM %{y:,.0f}<extra></extra>",
        ))
    fig_ov_monthly.update_layout(
        title="Monthly Revenue — AGENT1 vs AGENT2",
        plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG, barmode="group",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=sorted_months),
        yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
        margin=dict(t=50, b=80, l=90, r=20), height=340
    )

    # ── Overview: Credit Term by Agent ───────────────────────────────────────
    if "CREDITTERM" in df.columns:
        df_ov_c = df[df["Agent"].isin(IN_HOUSE) & df["CREDITTERM"].notna()].copy()
        df_ov_c["CreditTermNorm"] = df_ov_c["CREDITTERM"].apply(_norm_credit)
        df_ov_c = df_ov_c[df_ov_c["CreditTermNorm"].notna()]
        ov_credit = df_ov_c.groupby(["Agent","CreditTermNorm"]).agg(
            Revenue=("Revenue_RM","sum"), Customers=("Customer","nunique")
        ).reset_index()
        ov_credit["_sort"] = ov_credit["CreditTermNorm"].apply(_days_key)
        ov_credit = ov_credit.sort_values("_sort")
        all_terms = ov_credit["CreditTermNorm"].unique().tolist()

        fig_ov_credit = go.Figure()
        for agent in IN_HOUSE:
            sub = ov_credit[ov_credit["Agent"] == agent].set_index("CreditTermNorm")
            fig_ov_credit.add_trace(go.Bar(
                x=all_terms,
                y=[float(sub.loc[t,"Revenue"]) if t in sub.index else 0 for t in all_terms],
                name=agent, marker_color=AGENT_COLOR[agent],
                hovertemplate="%{x}<br>" + agent + ": RM %{y:,.0f}<extra></extra>",
            ))
        fig_ov_credit.update_layout(
            title="Revenue by Credit Term — AGENT1 vs AGENT2",
            plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG, barmode="group",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            xaxis=dict(type="category"),
            yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
            margin=dict(t=50, b=40, l=90, r=20), height=320
        )
        ov_credit_html = fig_html(fig_ov_credit)
    else:
        ov_credit_html = ""

    # ── Overview: Agent Summary Table ─────────────────────────────────────────
    agt_sum = df.groupby("Agent").agg(
        Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum"),
        Invoices=("Doc No","nunique"), Customers=("Customer","nunique"),
    ).reset_index()
    agt_sum["MarginPct"] = agt_sum["Margin"] / agt_sum["Revenue"] * 100
    agt_sum["AvgOrder"]  = agt_sum["Revenue"] / agt_sum["Invoices"]
    agt_sum["RevShare"]  = agt_sum["Revenue"] / rev * 100

    ov_rows = ""
    for _, r in agt_sum[agt_sum["Agent"].isin(IN_HOUSE)].iterrows():
        c = AGENT_COLOR.get(r["Agent"], "#333")
        ov_rows += (
            f'<tr><td style="padding:9px 12px;font-weight:700;color:{c}">{r["Agent"]}</td>'
            f'<td style="padding:9px 12px;text-align:right;font-weight:600">RM {r["Revenue"]:,.0f}</td>'
            f'<td style="padding:9px 12px;text-align:right">{r["RevShare"]:.1f}%</td>'
            f'<td style="padding:9px 12px;text-align:right">{r["MarginPct"]:.1f}%</td>'
            f'<td style="padding:9px 12px;text-align:right">{int(r["Invoices"]):,}</td>'
            f'<td style="padding:9px 12px;text-align:right">RM {r["AvgOrder"]:,.0f}</td>'
            f'<td style="padding:9px 12px;text-align:right">{int(r["Customers"])}</td></tr>'
        )
    ov_table = f"""
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead><tr style="background:#1a3c5e;color:white">
        <th style="padding:8px 12px;text-align:left">Agent</th>
        <th style="padding:8px 12px;text-align:right">Revenue</th>
        <th style="padding:8px 12px;text-align:right">% Share</th>
        <th style="padding:8px 12px;text-align:right">Margin %</th>
        <th style="padding:8px 12px;text-align:right">Invoices</th>
        <th style="padding:8px 12px;text-align:right">Avg Invoice</th>
        <th style="padding:8px 12px;text-align:right">Customers</th>
      </tr></thead>
      <tbody>{ov_rows}</tbody>
    </table>"""

    # ── Per-Agent Section ─────────────────────────────────────────────────────
    def _agent_section(agent):
        color    = AGENT_COLOR[agent]
        lcolor   = LOST_COLOR[agent]
        df_a     = df[df["Agent"] == agent].copy()
        full_a   = df_full[df_full["Agent"].astype(str) == agent].copy()
        if df_a.empty:
            return f'<div style="padding:16px;color:#aaa">{agent}: no data</div>'

        r_tot    = df_a["Revenue_RM"].sum()
        m_tot    = df_a["Margin_RM"].sum()
        mp       = m_tot / r_tot * 100 if r_tot else 0
        nm       = df_a["YearMonth"].nunique()
        ps       = df_a["Date"].min().strftime("%b %Y")
        pe       = df_a["Date"].max().strftime("%b %Y")

        # KPI cards
        cards_a = "".join([
            card_html("Revenue",         fmt_rm(r_tot),               f"{ps} – {pe}",    color),
            card_html("Gross Margin",    fmt_rm(m_tot),               f"{mp:.1f}%",      "#27ae60"),
            card_html("Avg / Month",     fmt_rm(r_tot / nm),          "per month",       "#2980b9"),
            card_html("Avg Margin / Month", fmt_rm(m_tot / nm),       f"{mp:.1f}% margin", "#27ae60"),
            card_html("Invoices",        f"{df_a['Doc No'].nunique():,}", "documents",    "#8e44ad"),
            card_html("Customers",       f"{df_a['Customer'].nunique():,}", "unique",     "#e67e22"),
        ])

        # Monthly trend
        m_trend = df_a.groupby(["YearMonth","MonthLabel"])["Revenue_RM"].sum().reset_index().sort_values("YearMonth")
        fig_trend = go.Figure(go.Scatter(
            x=m_trend["MonthLabel"].tolist(), y=m_trend["Revenue_RM"].tolist(),
            mode="lines+markers", line=dict(color=color, width=2.5), marker=dict(size=6),
            hovertemplate="%{x}<br>RM %{y:,.0f}<extra></extra>",
        ))
        fig_trend.update_layout(
            title=f"{agent} — Monthly Revenue Trend",
            plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
            xaxis=dict(tickangle=-45),
            yaxis=dict(tickprefix="RM ", tickformat=",.0f"),
            margin=dict(t=50, b=80, l=90, r=20), height=300
        )

        # Top 10 customers
        top_cust = df_a.groupby("Customer")["Revenue_RM"].sum().nlargest(10).reset_index().sort_values("Revenue_RM")
        fig_cust = go.Figure(go.Bar(
            x=top_cust["Revenue_RM"].tolist(), y=top_cust["Customer"].tolist(), orientation="h",
            marker_color=color,
            hovertemplate="%{y}<br>RM %{x:,.0f}<extra></extra>",
        ))
        fig_cust.update_layout(
            title=f"{agent} — Top 10 Customers",
            plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
            xaxis=dict(tickprefix="RM ", tickformat=",.0f"),
            margin=dict(t=50, b=20, l=220, r=20), height=360
        )

        # Top 15 products table
        top_prod = df_a.groupby("Product")["Revenue_RM"].sum().nlargest(15).reset_index()
        top_prod["Share"] = top_prod["Revenue_RM"] / r_tot * 100
        prod_rows = "".join(
            f'<tr style="background:{"#f8f9fa" if i%2==0 else CARD_BG}">'
            f'<td style="padding:6px 10px;text-align:center;color:#888">{i}</td>'
            f'<td style="padding:6px 10px">{r["Product"]}</td>'
            f'<td style="padding:6px 10px;text-align:right;font-weight:600">RM {r["Revenue_RM"]:,.0f}</td>'
            f'<td style="padding:6px 10px;text-align:right">{r["Share"]:.1f}%</td></tr>'
            for i, (_, r) in enumerate(top_prod.sort_values("Revenue_RM", ascending=False).iterrows(), 1)
        )
        prod_table = f"""
        <div style="font-size:14px;font-weight:700;color:#333;margin-bottom:8px">{agent} — Top 15 Products</div>
        <table style="width:100%;border-collapse:collapse;font-size:12px">
          <thead><tr style="background:{color};color:white">
            <th style="padding:7px 10px">#</th>
            <th style="padding:7px 10px;text-align:left">Product</th>
            <th style="padding:7px 10px;text-align:right">Revenue</th>
            <th style="padding:7px 10px;text-align:right">% Share</th>
          </tr></thead>
          <tbody>{prod_rows}</tbody>
        </table>"""

        # New vs Returning per month
        full_ym = full_a.copy()
        full_ym["YearMonth"] = full_ym["Date"].dt.to_period("M").astype(str)
        first_ever = full_ym.groupby("Customer")["YearMonth"].min()
        cm = df_a.groupby(["YearMonth","MonthLabel","Customer"]).size().reset_index(name="n")
        cm = cm.merge(first_ever.rename("FirstEver"), on="Customer", how="left")
        cm["Type"] = cm.apply(lambda r: "New" if r["YearMonth"] == r["FirstEver"] else "Returning", axis=1)
        rc = cm.groupby(["YearMonth","MonthLabel","Type"])["Customer"].nunique().reset_index().sort_values("YearMonth")
        mo = rc.drop_duplicates("YearMonth").sort_values("YearMonth")["MonthLabel"].tolist()
        fig_nr = go.Figure()
        for ctype, ccolor in [("Returning", "#2980b9"), ("New", "#27ae60")]:
            sub = rc[rc["Type"] == ctype].set_index("MonthLabel")
            fig_nr.add_trace(go.Bar(
                x=mo, y=[sub.loc[m, "Customer"] if m in sub.index else 0 for m in mo],
                name=ctype, marker_color=ccolor,
                hovertemplate="%{x}<br>" + ctype + ": %{y}<extra></extra>",
            ))
        fig_nr.update_layout(
            title=f"{agent} — New vs Returning Customers per Month",
            plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG, barmode="stack",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=mo),
            yaxis=dict(title="Customers"),
            margin=dict(t=50, b=80, l=60, r=20), height=300
        )

        # New / Lost customers per year
        display_years  = sorted(df_a["Year"].dropna().astype(int).unique())
        agt_first_year = full_a.groupby("Customer")["Year"].min()
        agt_years_all  = sorted(full_a["Year"].dropna().astype(int).unique())
        no2 = '<tr><td colspan="2" style="padding:6px;color:#aaa;text-align:center">—</td></tr>'

        flow_html = ""
        fig_flow  = go.Figure()
        new_counts, lost_counts = [], []
        for year in display_years:
            yr_a     = df_a[df_a["Year"] == year]
            nm_      = yr_a["Customer"].map(agt_first_year).astype("Int64") == year
            new_rev  = yr_a[nm_].groupby("Customer")["Revenue_RM"].sum().sort_values(ascending=False)
            prev     = year - 1
            if prev in agt_years_all:
                pc = set(full_a[full_a["Year"] == prev]["Customer"].unique())
                cc = set(full_a[full_a["Year"] == year]["Customer"].unique())
                lost_rev = full_a[(full_a["Year"] == prev) & full_a["Customer"].isin(pc - cc)].groupby("Customer")["Revenue_RM"].sum().sort_values(ascending=False)
            else:
                lost_rev = pd.Series(dtype=float)
            new_counts.append(len(new_rev))
            lost_counts.append(len(lost_rev))

            new_rows  = "".join(f'<tr style="background:{"#f8f9fa" if i%2==0 else CARD_BG}"><td style="padding:5px 8px">{c}</td><td style="padding:5px 8px;text-align:right;color:#27ae60;font-weight:600">RM {v:,.0f}</td></tr>' for i,(c,v) in enumerate(new_rev.items(),1))
            lost_rows = "".join(f'<tr style="background:{"#f8f9fa" if i%2==0 else CARD_BG}"><td style="padding:5px 8px">{c}</td><td style="padding:5px 8px;text-align:right;color:{lcolor};font-weight:600">RM {v:,.0f}</td></tr>' for i,(c,v) in enumerate(lost_rev.items(),1))
            flow_html += f"""
            <div style="margin-bottom:20px">
              <div style="font-size:13px;font-weight:700;color:#1a3c5e;border-bottom:1px solid #ddd;padding-bottom:4px;margin-bottom:10px">{year}</div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
                <div>
                  <div style="font-size:12px;font-weight:700;color:#27ae60;margin-bottom:4px">▲ New ({len(new_rev)})</div>
                  <table style="width:100%;border-collapse:collapse;font-size:11px">
                    <thead><tr style="background:#27ae60;color:white"><th style="padding:5px 8px;text-align:left">Customer</th><th style="padding:5px 8px;text-align:right">Revenue {year}</th></tr></thead>
                    <tbody>{new_rows if new_rows else no2}</tbody>
                  </table>
                </div>
                <div>
                  <div style="font-size:12px;font-weight:700;color:{lcolor};margin-bottom:4px">▼ Lost from {year-1} ({len(lost_rev)})</div>
                  <table style="width:100%;border-collapse:collapse;font-size:11px">
                    <thead><tr style="background:{lcolor};color:white"><th style="padding:5px 8px;text-align:left">Customer</th><th style="padding:5px 8px;text-align:right">Last Revenue ({year-1})</th></tr></thead>
                    <tbody>{lost_rows if lost_rows else no2}</tbody>
                  </table>
                </div>
              </div>
            </div>"""

        fig_flow.add_trace(go.Bar(x=[str(y) for y in display_years], y=new_counts,  name="New",  marker_color="#27ae60", hovertemplate="%{x} New: %{y}<extra></extra>"))
        fig_flow.add_trace(go.Bar(x=[str(y) for y in display_years], y=[-v for v in lost_counts], name="Lost", marker_color=lcolor, customdata=lost_counts, hovertemplate="%{x} Lost: %{customdata}<extra></extra>"))
        net = [n - l for n, l in zip(new_counts, lost_counts)]
        fig_flow.add_trace(go.Scatter(x=[str(y) for y in display_years], y=net, name="Net", mode="lines+markers+text", text=[f"+{n}" if n>=0 else str(n) for n in net], textposition="top center", line=dict(color="#1a3c5e", width=2), marker=dict(size=7), hovertemplate="%{x} Net: %{y}<extra></extra>"))
        fig_flow.update_layout(
            title=f"{agent} — Customer Acquisition & Churn",
            plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG, barmode="relative",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            yaxis=dict(title="Customers", zeroline=True, zerolinecolor="#ccc"),
            xaxis=dict(type="category"),
            margin=dict(t=50, b=40, l=60, r=20), height=320
        )

        # Reorder prediction
        inv_d = df_a.groupby(["Customer","Doc No"])["Date"].min().reset_index().sort_values(["Customer","Date"])
        last_date = df_a["Date"].max()
        freq_rows_a = []
        for cust, grp in inv_d.groupby("Customer"):
            dates = sorted(grp["Date"].tolist())
            if len(dates) >= 5:
                gaps = [(dates[i+1]-dates[i]).days for i in range(len(dates)-1)]
                avg_gap = round(sum(gaps)/len(gaps))
                if avg_gap >= 10:
                    predicted = dates[-1] + pd.Timedelta(days=avg_gap)
                    freq_rows_a.append({"Customer": cust, "Orders": len(dates), "LastOrder": dates[-1], "AvgGapDays": avg_gap, "PredictedNext": predicted, "DaysUntil": (predicted - last_date).days})
        freq_a = pd.DataFrame(freq_rows_a)
        due_a   = freq_a[freq_a["DaysUntil"] <= 30].sort_values("DaysUntil") if not freq_a.empty else pd.DataFrame()
        watch_a = freq_a[(freq_a["DaysUntil"] > 30) & (freq_a["DaysUntil"] <= 90)].sort_values("DaysUntil") if not freq_a.empty else pd.DataFrame()

        def _freq_tbl(title, df_sub, hcolor):
            no_r = '<tr><td colspan="5" style="padding:8px;color:#aaa;text-align:center">—</td></tr>'
            rows = ""
            for i, (_, r) in enumerate(df_sub.iterrows(), 1):
                d = int(r["DaysUntil"])
                bc = "#e74c3c" if d < 0 else ("#e67e22" if d == 0 else "#27ae60")
                bt = f"Overdue {abs(d)}d" if d < 0 else ("Today" if d == 0 else f"In {d}d")
                rows += (f'<tr style="background:{"#f8f9fa" if i%2==0 else CARD_BG}">'
                         f'<td style="padding:6px 8px">{r["Customer"]}</td>'
                         f'<td style="padding:6px 8px;text-align:center">{int(r["Orders"])}</td>'
                         f'<td style="padding:6px 8px;text-align:center">{r["LastOrder"].strftime("%d %b %Y")}</td>'
                         f'<td style="padding:6px 8px;text-align:center">{int(r["AvgGapDays"])}d</td>'
                         f'<td style="padding:6px 8px;text-align:center"><span style="background:{bc};color:white;padding:2px 7px;border-radius:10px;font-size:11px;font-weight:600">{bt}</span></td></tr>')
            return f"""<div style="margin-bottom:16px">
              <div style="font-size:12px;font-weight:700;color:{hcolor};margin-bottom:4px">{title}</div>
              <table style="width:100%;border-collapse:collapse;font-size:11px">
                <thead><tr style="background:{hcolor};color:white">
                  <th style="padding:6px 8px;text-align:left">Customer</th><th style="padding:6px 8px">Orders</th>
                  <th style="padding:6px 8px">Last Order</th><th style="padding:6px 8px">Avg Freq</th><th style="padding:6px 8px">Status</th>
                </tr></thead>
                <tbody>{rows if rows else no_r}</tbody>
              </table></div>"""

        reorder_a = (
            _freq_tbl(f"Due within 30 days ({len(due_a)} customers)", due_a, "#e74c3c") +
            _freq_tbl(f"Coming up 31–90 days ({len(watch_a)} customers)", watch_a, "#e67e22")
        )

        # ── Margin Trend ──────────────────────────────────────────────────────
        mgn_mo = df_a.groupby(["YearMonth","MonthLabel"]).agg(
            Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum")
        ).reset_index().sort_values("YearMonth")
        mgn_mo["MarginPct"] = (mgn_mo["Margin"] / mgn_mo["Revenue"] * 100).round(1)
        avg_mp = float(mgn_mo["MarginPct"].mean())
        mo_labels = mgn_mo["MonthLabel"].tolist()

        fig_mgn_a = go.Figure()
        fig_mgn_a.add_trace(go.Bar(
            x=mo_labels, y=mgn_mo["Margin"].tolist(),
            name="Margin (RM)", marker_color=color, opacity=0.5,
            hovertemplate="%{x}<br>Margin: RM %{y:,.0f}<extra></extra>", yaxis="y1",
        ))
        fig_mgn_a.add_trace(go.Scatter(
            x=mo_labels, y=mgn_mo["MarginPct"].tolist(),
            name="Margin %", mode="lines+markers+text",
            text=[f"{v:.1f}%" for v in mgn_mo["MarginPct"]],
            textposition="top center", textfont=dict(size=9),
            line=dict(color="#1a3c5e", width=2.5), marker=dict(size=5),
            hovertemplate="%{x}<br>Margin %%: %{y:.1f}%%<extra></extra>", yaxis="y2",
        ))
        fig_mgn_a.add_trace(go.Scatter(
            x=mo_labels, y=[avg_mp]*len(mo_labels),
            name=f"Avg {avg_mp:.1f}%", mode="lines",
            line=dict(color="#e74c3c", dash="dash", width=1.5),
            hovertemplate=f"Avg margin: {avg_mp:.1f}%<extra></extra>", yaxis="y2",
        ))
        fig_mgn_a.update_layout(
            title=f"{agent} — Monthly Gross Margin Trend",
            plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            xaxis=dict(tickangle=-45),
            yaxis=dict(tickprefix="RM ", tickformat=",.0f", title="Margin RM"),
            yaxis2=dict(ticksuffix="%", title="Margin %", overlaying="y", side="right",
                        showgrid=False, range=[0, max(mgn_mo["MarginPct"])*1.3]),
            margin=dict(t=50, b=80, l=90, r=70), height=340
        )

        # ── AOV Trend ─────────────────────────────────────────────────────────
        aov_mo = df_a.groupby(["YearMonth","MonthLabel"]).agg(
            Revenue=("Revenue_RM","sum"), Invoices=("Doc No","nunique")
        ).reset_index().sort_values("YearMonth")
        aov_mo["AOV"] = aov_mo["Revenue"] / aov_mo["Invoices"]
        avg_aov_a = float(aov_mo["AOV"].mean())

        fig_aov_a = go.Figure()
        fig_aov_a.add_trace(go.Bar(
            x=aov_mo["MonthLabel"].tolist(), y=aov_mo["Invoices"].tolist(),
            name="Invoices", marker_color="#bdc3c7", opacity=0.4,
            hovertemplate="%{x}<br>Invoices: %{y}<extra></extra>", yaxis="y2",
        ))
        fig_aov_a.add_trace(go.Scatter(
            x=aov_mo["MonthLabel"].tolist(), y=aov_mo["AOV"].tolist(),
            name="AOV", mode="lines+markers+text",
            text=[f"RM {v:,.0f}" for v in aov_mo["AOV"]],
            textposition="top center", textfont=dict(size=9),
            line=dict(color=color, width=2.5), marker=dict(size=5),
            hovertemplate="%{x}<br>AOV: RM %{y:,.0f}<extra></extra>", yaxis="y1",
        ))
        fig_aov_a.add_trace(go.Scatter(
            x=aov_mo["MonthLabel"].tolist(), y=[avg_aov_a]*len(aov_mo),
            name=f"Avg RM {avg_aov_a:,.0f}", mode="lines",
            line=dict(color="#e74c3c", dash="dash", width=1.5),
            hovertemplate=f"Avg AOV: RM {avg_aov_a:,.0f}<extra></extra>", yaxis="y1",
        ))
        fig_aov_a.update_layout(
            title=f"{agent} — Average Order Value (AOV) per Month",
            plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            xaxis=dict(tickangle=-45),
            yaxis=dict(tickprefix="RM ", tickformat=",.0f", title="AOV (RM)"),
            yaxis2=dict(title="Invoices", overlaying="y", side="right", showgrid=False),
            margin=dict(t=50, b=80, l=90, r=70), height=320
        )

        # ── Brand Breakdown ───────────────────────────────────────────────────
        brand_a = df_a.groupby("Brand").agg(
            Revenue=("Revenue_RM","sum"), Margin=("Margin_RM","sum")
        ).reset_index()
        brand_a = brand_a[brand_a["Revenue"] > 0].copy()
        brand_a["MarginPct"] = (brand_a["Margin"] / brand_a["Revenue"] * 100).round(1)
        brand_a["Share"]     = (brand_a["Revenue"] / brand_a["Revenue"].sum() * 100).round(1)
        brand_a = brand_a.sort_values("Revenue", ascending=False)

        brand_rows_a = ""
        for i, (_, br) in enumerate(brand_a.iterrows(), 1):
            bg = "#f8f9fa" if i % 2 == 0 else CARD_BG
            mc = "#27ae60" if br["MarginPct"] >= 45 else "#e74c3c"
            brand_rows_a += (
                f'<tr style="background:{bg}">'
                f'<td style="padding:6px 10px">{br["Brand"]}</td>'
                f'<td style="padding:6px 10px;text-align:right;font-weight:600">RM {br["Revenue"]:,.0f}</td>'
                f'<td style="padding:6px 10px;text-align:right">{br["Share"]:.1f}%</td>'
                f'<td style="padding:6px 10px;text-align:right;color:{mc};font-weight:600">{br["MarginPct"]:.1f}%</td>'
                f'</tr>'
            )
        brand_table_a = f"""
        <div style="font-size:13px;font-weight:700;color:#333;margin-bottom:8px">{agent} — Brand Breakdown</div>
        <table style="width:100%;border-collapse:collapse;font-size:12px">
          <thead><tr style="background:{color};color:white">
            <th style="padding:6px 10px;text-align:left">Brand</th>
            <th style="padding:6px 10px;text-align:right">Revenue</th>
            <th style="padding:6px 10px;text-align:right">% Share</th>
            <th style="padding:6px 10px;text-align:right">Margin %</th>
          </tr></thead>
          <tbody>{brand_rows_a}</tbody>
        </table>"""

        # ── 80/20 Customer Pareto ─────────────────────────────────────────────
        cp_a = df_a.groupby("Customer")["Revenue_RM"].sum().sort_values(ascending=False).reset_index()
        cp_a["CumPct"]  = cp_a["Revenue_RM"].cumsum() / cp_a["Revenue_RM"].sum() * 100
        cp_a["CustPct"] = (cp_a.index + 1) / len(cp_a) * 100
        n80_a   = len(cp_a[cp_a["CumPct"] <= 80]) + 1
        pct80_a = round(n80_a / len(cp_a) * 100, 1)

        fig_pareto_a = go.Figure()
        fig_pareto_a.add_trace(go.Scatter(
            x=cp_a["CustPct"].tolist(), y=cp_a["CumPct"].tolist(),
            mode="lines", line=dict(color=color, width=2.5),
            hovertemplate="Top %{x:.1f}% customers → %{y:.1f}% revenue<extra></extra>",
        ))
        fig_pareto_a.add_hline(y=80, line_dash="dash", line_color="#e74c3c",
            annotation_text="80% revenue", annotation_position="top left")
        fig_pareto_a.add_vline(x=pct80_a, line_dash="dash", line_color="#e67e22",
            annotation_text=f"{pct80_a}% of customers", annotation_position="top right")
        fig_pareto_a.update_layout(
            title=f"{agent} — Customer Pareto: {pct80_a}% of customers = 80% revenue ({n80_a}/{len(cp_a)})",
            plot_bgcolor=CARD_BG, paper_bgcolor=CARD_BG,
            xaxis=dict(ticksuffix="%", title="Cumulative % of Customers"),
            yaxis=dict(ticksuffix="%", title="Cumulative % of Revenue"),
            margin=dict(t=60, b=50, l=70, r=20), height=320
        )

        # ── RFM Segmentation ──────────────────────────────────────────────────
        rfm_a = df_a.groupby("Customer").agg(
            LastOrder =("Date",       "max"),
            Frequency =("Doc No",     "nunique"),
            Monetary  =("Revenue_RM", "sum"),
        ).reset_index()
        rfm_a["Recency"] = (df_a["Date"].max() - rfm_a["LastOrder"]).dt.days

        if len(rfm_a) >= 5:
            rfm_a["R"] = pd.qcut(rfm_a["Recency"].rank(method="first"),  5, labels=[5,4,3,2,1]).astype(int)
            rfm_a["F"] = pd.qcut(rfm_a["Frequency"].rank(method="first"),5, labels=[1,2,3,4,5]).astype(int)
            rfm_a["M"] = pd.qcut(rfm_a["Monetary"].rank(method="first"), 5, labels=[1,2,3,4,5]).astype(int)
        else:
            for col in ["R","F","M"]: rfm_a[col] = 3

        SEG_META_A = {
            "🏆 Champion":    {"color":"#27ae60","bg":"#eafaf1","action":"VIP 维护，优先推介新品"},
            "💚 Loyal":       {"color":"#2980b9","bg":"#eaf4fb","action":"定期回访，交叉销售"},
            "🌱 Promising":   {"color":"#16a085","bg":"#e8f8f5","action":"新客培养，扩大覆盖"},
            "💛 Potential":   {"color":"#d4ac0d","bg":"#fef9e7","action":"提升购买频率"},
            "⚠️ At Risk":     {"color":"#e67e22","bg":"#fef5ec","action":"紧急跟进"},
            "😴 Hibernating": {"color":"#7f8c8d","bg":"#f2f3f4","action":"季度性激活"},
            "❌ Lost":        {"color":"#e74c3c","bg":"#fdedec","action":"考虑重新激活"},
        }

        def _seg(r, f, m):
            if r>=5 and f>=4 and m>=4: return "🏆 Champion"
            if r>=4 and f>=3 and m>=3: return "💚 Loyal"
            if r>=4 and f<=2:          return "🌱 Promising"
            if r>=3 and f>=2 and m>=2: return "💛 Potential"
            if r<=2 and f>=3:          return "⚠️ At Risk"
            if r<=2 and f<=2 and m<=2: return "😴 Hibernating"
            return "❌ Lost"

        rfm_a["Segment"] = rfm_a.apply(lambda row: _seg(row["R"],row["F"],row["M"]), axis=1)

        # RFM — collapsible dropdown per segment (same style as main dashboard)
        seg_order_a = list(SEG_META_A.keys())
        rfm_detail_a = f'<div style="font-size:13px;font-weight:700;color:#333;margin-bottom:10px">{agent} — RFM Customer Segmentation</div>'
        for seg in seg_order_a:
            meta = SEG_META_A[seg]
            sub  = rfm_a[rfm_a["Segment"] == seg].sort_values("Monetary", ascending=False)
            if sub.empty: continue
            rows_rfm = ""
            for i, (_, r) in enumerate(sub.iterrows(), 1):
                bg = "#f8f9fa" if i%2==0 else CARD_BG
                rows_rfm += (
                    f'<tr style="background:{bg}">'
                    f'<td style="padding:6px 10px">{r["Customer"]}</td>'
                    f'<td style="padding:6px 10px;text-align:center">{int(r["Recency"])}d</td>'
                    f'<td style="padding:6px 10px;text-align:center">{int(r["Frequency"])}</td>'
                    f'<td style="padding:6px 10px;text-align:right;font-weight:600">RM {r["Monetary"]:,.0f}</td>'
                    f'<td style="padding:6px 10px;text-align:center">{r["R"]}-{r["F"]}-{r["M"]}</td>'
                    f'</tr>'
                )
            total_rev = sub["Monetary"].sum()
            rfm_detail_a += f"""
        <details style="margin-bottom:10px">
          <summary style="cursor:pointer;padding:10px 14px;background:{meta.get('bg','#f5f5f5')};
                          border-left:4px solid {meta['color']};border-radius:6px;
                          font-weight:700;color:{meta['color']};list-style:none;
                          display:flex;justify-content:space-between;align-items:center">
            <span>{seg} &nbsp;·&nbsp; {len(sub)} customers &nbsp;·&nbsp; RM {total_rev:,.0f}</span>
            <span style="font-size:11px;color:#888;font-weight:400">{meta['action']}</span>
          </summary>
          <div style="padding:0 4px">
            <table style="width:100%;border-collapse:collapse;font-size:12px">
              <thead><tr style="background:{meta['color']};color:white">
                <th style="padding:6px 10px;text-align:left">Customer</th>
                <th style="padding:6px 10px">Days Since Order</th>
                <th style="padding:6px 10px">Invoices</th>
                <th style="padding:6px 10px;text-align:right">Total Revenue</th>
                <th style="padding:6px 10px">R-F-M</th>
              </tr></thead>
              <tbody>{rows_rfm}</tbody>
            </table>
          </div>
        </details>"""
        rfm_table_a = rfm_detail_a

        # ── Revenue at Risk ───────────────────────────────────────────────────
        lo_a   = df_a.groupby("Customer")["Date"].max()
        di_a   = (df_a["Date"].max() - lo_a).dt.days
        r12_a  = df_a[df_a["Date"] >= df_a["Date"].max() - pd.DateOffset(months=12)].groupby("Customer")["Revenue_RM"].sum()
        risk_a = pd.DataFrame({
            "Customer":   di_a.index,
            "DaysSince":  di_a.values,
            "LastOrder":  lo_a.values,
            "Revenue12m": [float(r12_a.get(c,0)) for c in di_a.index],
        })
        risk_a = risk_a[risk_a["DaysSince"] >= 90].sort_values("DaysSince", ascending=False)

        n_critical_a = len(risk_a[risk_a["DaysSince"] >= 180])
        n_watch_a    = len(risk_a[(risk_a["DaysSince"] >= 90) & (risk_a["DaysSince"] < 180)])
        rev_critical_a = float(risk_a[risk_a["DaysSince"] >= 180]["Revenue12m"].sum())
        rev_watch_a    = float(risk_a[(risk_a["DaysSince"] >= 90) & (risk_a["DaysSince"] < 180)]["Revenue12m"].sum())

        risk_cards_a = "".join([
            card_html("Inactive 90d+",    f"{len(risk_a)}",          "customers need follow-up", "#e74c3c"),
            card_html("Revenue at Risk",  fmt_rm(risk_a["Revenue12m"].sum()), "from inactive customers", "#e67e22"),
            card_html("Critical (180d+)", fmt_rm(rev_critical_a),   f"{n_critical_a} customers", "#e74c3c"),
            card_html("Watch (90–179d)",  fmt_rm(rev_watch_a),      f"{n_watch_a} customers",    "#e67e22"),
        ])

        risk_rows_a = ""
        for i, (_, r) in enumerate(risk_a.iterrows(), 1):
            label, rc = ("Critical (180d+)", "#e74c3c") if r["DaysSince"]>=180 else ("Watch (90–179d)", "#e67e22")
            bg = "#f8f9fa" if i%2==0 else CARD_BG
            risk_rows_a += (
                f'<tr style="background:{bg}">'
                f'<td style="padding:6px 10px">{r["Customer"]}</td>'
                f'<td style="padding:6px 10px;text-align:center">{r["LastOrder"].strftime("%d %b %Y")}</td>'
                f'<td style="padding:6px 10px;text-align:center">{int(r["DaysSince"])}d</td>'
                f'<td style="padding:6px 10px;text-align:right;font-weight:600">RM {r["Revenue12m"]:,.0f}</td>'
                f'<td style="padding:6px 10px;text-align:center">'
                f'<span style="background:{rc};color:white;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">{label}</span>'
                f'</td></tr>'
            )
        no_risk_a = '<tr><td colspan="5" style="padding:10px;text-align:center;color:#aaa">No inactive customers</td></tr>'
        risk_tbl_a = f"""
        <div style="font-size:14px;font-weight:700;color:#1a3c5e;margin-bottom:12px">{agent} — Revenue at Risk</div>
        <div class="cards" style="margin-bottom:16px">{risk_cards_a}</div>
        <table style="width:100%;border-collapse:collapse;font-size:12px">
          <thead><tr style="background:#1a3c5e;color:white">
            <th style="padding:7px 10px;text-align:left">Customer</th>
            <th style="padding:7px 10px">Last Order</th>
            <th style="padding:7px 10px">Days Inactive</th>
            <th style="padding:7px 10px;text-align:right">Revenue (12m)</th>
            <th style="padding:7px 10px">Status</th>
          </tr></thead>
          <tbody>{risk_rows_a if risk_rows_a else no_risk_a}</tbody>
        </table>"""

        return f"""
        <div style="background:#f0f4f8;border-left:5px solid {color};padding:12px 20px;margin-bottom:16px;border-radius:4px">
          <div style="font-size:18px;font-weight:700;color:{color}">{agent}</div>
          <div style="font-size:12px;color:#888">{ps} – {pe}</div>
        </div>
        <div class="cards">{cards_a}</div>
        <div class="row row-2">
          <div class="panel">{fig_html(fig_trend)}</div>
          <div class="panel">{fig_html(fig_nr)}</div>
        </div>
        <div class="row row-2">
          <div class="panel">{fig_html(fig_mgn_a)}</div>
          <div class="panel">{fig_html(fig_aov_a)}</div>
        </div>
        <div class="row row-1"><div class="panel-p">{brand_table_a}</div></div>
        <div class="row row-1"><div class="panel">{fig_html(fig_pareto_a)}</div></div>
        <div class="row row-1"><div class="panel-p">{rfm_table_a}</div></div>
        <div class="row row-1"><div class="panel-p">{risk_tbl_a}</div></div>
        <div class="row row-1"><div class="panel">{fig_html(fig_cust)}</div></div>
        <div class="row row-1"><div class="panel-p">{prod_table}</div></div>
        <div class="row row-1"><div class="panel">{fig_html(fig_flow)}</div></div>
        <div class="row row-1"><div class="panel-p">{flow_html}</div></div>
        <div class="row row-1"><div class="panel-p">{reorder_a}</div></div>"""

    agent_sections_html = ""
    for agt in IN_HOUSE:
        agent_sections_html += f"""
  <hr class="divider">
  {section_header(agt + " — Individual Performance")}
  {_agent_section(agt)}"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>LLG Industries — {" & ".join(IN_HOUSE)} Sales Report</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:'Segoe UI',sans-serif; background:{BG}; color:#333; }}
  .header {{ background:#1a3c5e; color:white; padding:24px 32px; }}
  .header h1 {{ font-size:22px; font-weight:700; }}
  .header p  {{ font-size:13px; opacity:.7; margin-top:4px; }}
  .container {{ max-width:1400px; margin:0 auto; padding:24px; }}
  .cards  {{ display:flex; gap:12px; margin-bottom:16px; flex-wrap:wrap; }}
  .row    {{ display:grid; gap:16px; margin-bottom:16px; }}
  .row-2  {{ grid-template-columns:1fr 1fr; }}
  .row-1  {{ grid-template-columns:1fr; }}
  .panel   {{ background:{CARD_BG}; border-radius:12px; padding:8px; box-shadow:0 2px 8px rgba(0,0,0,0.08); }}
  .panel-p {{ background:{CARD_BG}; border-radius:12px; padding:16px; box-shadow:0 2px 8px rgba(0,0,0,0.08); }}
  .divider {{ border:none; border-top:3px solid #1a3c5e; margin:32px 0; }}
  .tbl-wrap {{ overflow-x:auto; -webkit-overflow-scrolling:touch; }}
  @media (max-width:768px) {{
    .header {{ padding:16px; }}
    .header h1 {{ font-size:17px; }}
    .container {{ padding:8px; overflow-x:hidden; }}
    .row-2 {{ grid-template-columns:1fr; }}
    .panel, .panel-p {{ padding:6px; }}
    .cards > div {{ min-width:calc(50% - 6px) !important; padding:14px 10px !important; }}
    .cards > div > div:nth-child(2) {{ font-size:20px !important; }}
    table {{ font-size:11px; display:block; overflow-x:auto; -webkit-overflow-scrolling:touch; max-width:100%; }}
    details summary {{ font-size:13px; }}
  }}
  @media (max-width:480px) {{
    .cards > div {{ min-width:100% !important; }}
    .cards > div > div:nth-child(2) {{ font-size:18px !important; }}
  }}
</style>
</head>
<body>
<div class="header">
  <h1>LLG Industries — {" & ".join(IN_HOUSE)} Sales Report</h1>
  <p>{" &amp; ".join(IN_HOUSE)} &nbsp;|&nbsp; {period_start} – {period_end} &nbsp;|&nbsp; Generated {generated}</p>
</div>
<div class="container">

  {section_header("Overview — " + " & ".join(IN_HOUSE) + " Combined" if len(IN_HOUSE) > 1 else "Overview — " + IN_HOUSE[0])}
  <div class="cards">{overview_cards}</div>
  <div class="row row-1"><div class="panel-p">{ov_table}</div></div>
  <div class="row row-1"><div class="panel">{fig_html(fig_ov_monthly)}</div></div>
  {f'<div class="row row-1"><div class="panel">{ov_credit_html}</div></div>' if ov_credit_html else ''}

  {agent_sections_html}

</div>
</body>
</html>"""

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    DATA_DIR = Path(__file__).parent.parent / "01_Raw Data" / "Apr26"
    FILES = {
        "LLE": str(DATA_DIR / "LLE SALES BY PROFIT & LOSS BY DOCUMENT 260430-230101.xlsx"),
        "LLG": str(DATA_DIR / "LLG SALES BY PROFIT & LOSS BY DOCUMENT 260430-230101.xlsx"),
    }

    print(f"\n{'='*55}")
    print(f"  LLG Industries — Dashboard Generator")
    print(f"{'='*55}")

    frames, foc_frames, full_frames = [], [], []
    for company, filename in FILES.items():
        if not Path(filename).exists():
            print(f"  Skipping: {filename}")
            continue
        print(f"  Loading {company}...", end="", flush=True)
        frames.append(load(filename, company))
        foc_frames.append(load_foc(filename, company))
        full_frames.append(load_full(filename))
        print(" done")

    if not frames:
        print("No data files found.")
        sys.exit(1)

    combined = pd.concat(frames, ignore_index=True)
    foc_df   = pd.concat(foc_frames, ignore_index=True) if foc_frames else pd.DataFrame(columns=["Cost_RM"])
    foc_cost = float(foc_df["Cost_RM"].sum()) if len(foc_df) else 0.0
    df_full  = pd.concat(full_frames, ignore_index=True) if full_frames else pd.DataFrame(columns=["Year","Customer","Revenue_RM"])

    print(f"  Rows: {len(combined):,} | Invoices: {combined['Doc No'].nunique():,} | FOC cost: RM {foc_cost:,.0f}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.now().strftime('%Y-%m-%d')

    # ── Report 1: Company Overview ────────────────────────────────────────────
    print(f"  Building company dashboard...", end="", flush=True)
    html = build_dashboard(combined, foc_cost, df_full,
                           report_title="LLG Industries Sdn Bhd — Sales Dashboard")
    print(" done")
    out_path = OUTPUT_DIR / f"dashboard_{today}.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"  {out_path.resolve()}")

    # ── Report 2: In-House Agent (AGENT1 & AGENT2 only) ──────────────────────
    IN_HOUSE = ["AGENT1", "AGENT2"]
    agent_df      = combined[combined["Agent"].isin(IN_HOUSE)].copy()
    agent_full_df = df_full[df_full["Agent"].astype(str).isin(IN_HOUSE)].copy()
    agent_foc_df  = pd.concat(
        [load_foc(filename, company) for company, filename in FILES.items() if Path(filename).exists()],
        ignore_index=True
    )
    agent_foc_df  = agent_foc_df[agent_foc_df["Customer"].isin(agent_df["Customer"].unique())]
    agent_foc_cost = float(agent_foc_df["Cost_RM"].sum()) if len(agent_foc_df) else 0.0

    if not agent_df.empty:
        print(f"  Building agent dashboard...", end="", flush=True)
        html_agent = build_agent_dashboard(agent_df, agent_foc_cost, agent_full_df)
        print(" done")
        agent_path = OUTPUT_DIR / f"dashboard_agent_{today}.html"
        agent_path.write_text(html_agent, encoding="utf-8")
        print(f"  {agent_path.resolve()}")

    # ── Report 3: Agent 3 (Boss) ──────────────────────────────────────────────
    AGENT3_LIST = ["AGENT3"]
    agent3_df      = combined[combined["Agent"].isin(AGENT3_LIST)].copy()
    agent3_full_df = df_full[df_full["Agent"].astype(str).isin(AGENT3_LIST)].copy()
    agent3_foc_df  = pd.concat(
        [load_foc(filename, company) for company, filename in FILES.items() if Path(filename).exists()],
        ignore_index=True
    )
    agent3_foc_df  = agent3_foc_df[agent3_foc_df["Customer"].isin(agent3_df["Customer"].unique())]
    agent3_foc_cost = float(agent3_foc_df["Cost_RM"].sum()) if len(agent3_foc_df) else 0.0

    if not agent3_df.empty:
        print(f"  Building agent3 dashboard...", end="", flush=True)
        html_agent3 = build_agent_dashboard(agent3_df, agent3_foc_cost, agent3_full_df, agents=AGENT3_LIST)
        print(" done")
        agent3_path = OUTPUT_DIR / f"dashboard_agent3_{today}.html"
        agent3_path.write_text(html_agent3, encoding="utf-8")
        print(f"  {agent3_path.resolve()}")
    else:
        print(f"  No AGENT3 data found — skipping agent3 dashboard")

    print(f"\n{'='*55}")
    print(f"  All reports ready!")
    print(f"{'='*55}\n")

if __name__ == "__main__":
    main()
