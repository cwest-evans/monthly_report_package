# report_pdf.py
from __future__ import annotations

import os
import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Any, Optional

import re
from typing import Tuple

import pandas as pd
import matplotlib.pyplot as plt

try:
    import seaborn as sns
    _HAS_SEABORN = True
except ImportError:
    sns = None
    _HAS_SEABORN = False

from jinja2 import Environment, FileSystemLoader, select_autoescape
from playwright.sync_api import sync_playwright

MONEY_COL_BY_SHEET = {
    "Job Revenue (Month)": "EarnedRevenue_Est",
    "Job Revenue (YTD)":   "EarnedRevenue_Est",
}
JOB_COL_BY_SHEET = {
    "Job Revenue (Month)": "Job",
    "Job Revenue (YTD)":   "Job",
}

BACKLOG_CHART_PERIODS = 18        # for chart (Rolling 18 only)
BACKLOG_TABLE_PERIODS = 6         # for table (condensed view)
BACKLOG_TABLE_TOP_N = 20          # number of jobs shown in backlog table


def drop_cols_if_exist(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    cols = [c for c in cols if c in df.columns]
    return df.drop(columns=cols) if cols else df


def rename_cols_if_exist(df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    mapping2 = {k: v for k, v in mapping.items() if k in df.columns}
    return df.rename(columns=mapping2) if mapping2 else df


def is_year_label(x: str) -> bool:
    s = str(x).strip()
    return bool(re.fullmatch(r"\d{4}", s))


def coerce_numeric_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


def pick_period_columns(backlog_df: pd.DataFrame) -> list[str]:
    """
    Backlog table shape: first 3 columns are metadata (JobSearch, Division, MarketSegment),
    then dynamic period label columns. We want the Rolling 18 period columns (exclude year rollups).
    """
    if backlog_df is None or backlog_df.empty or backlog_df.shape[1] <= 3:
        return []

    period_cols = list(backlog_df.columns[3:])
    period_cols_non_year = [c for c in period_cols if not is_year_label(str(c))]
    return period_cols_non_year


def prepare_rev_preview(df: pd.DataFrame) -> pd.DataFrame:
    """
    A) Top Jobs tables:
      1 Remove RunId
      2 Remove Job
      3 Rename JobSearch -> Job
      4 Remove FeeCollected_Est
    """
    if df is None or df.empty:
        return df

    dfx = df.copy()

    # Drop by name if present (preferred)
    dfx = drop_cols_if_exist(dfx, ["RunId", "Job", "FeeCollected_Est"])

    # Rename JobSearch -> Job
    dfx = rename_cols_if_exist(dfx, {"JobSearch": "Job"})

    # If the schema ever changes and those columns don't exist,
    # you *could* fallback to positional drops — but I recommend failing loudly instead.
    return dfx


def prepare_wip_preview(df: pd.DataFrame) -> pd.DataFrame:
    """
    B) WIP (Preview) – only show selected columns and rename JobName -> Job
    """
    if df is None or df.empty:
        return df

    wanted = [
        "JobName",
        "Division",
        "MarketSegment",
        "total_contract_amount",
        "cost_percent_complete",
        "earned_revenue",
        "billings_to_complete",
        "profit_percentage",
        "prev_fee_percentage",
        "percent_change",
    ]

    dfx = df.copy()
    available = [c for c in wanted if c in dfx.columns]

    # If your WIP uses different casing/aliases, fix it here once.
    dfx = dfx[available] if available else dfx

    dfx = rename_cols_if_exist(dfx, {"JobName": "Job"})
    return dfx


def prepare_backlog_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    D) Backlog Table – condensed:
      - Keep Job/Division/Market
      - Add Total_R18
      - Show first N rolling period columns (e.g., next 6 periods)
      - Keep top jobs by Total_R18
    """
    if df is None or df.empty or df.shape[1] <= 3:
        return df

    dfx = df.copy()

    meta_cols = list(dfx.columns[:3])  # typically JobSearch, Division, MarketSegment
    period_cols = pick_period_columns(dfx)
    if not period_cols:
        return dfx[meta_cols].copy()

    # Compute Total_R18 across non-year period cols
    numeric = dfx[period_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    dfx["Total_R18"] = numeric.sum(axis=1)

    # Choose a small set of period columns to display (first 6 is fine)
    show_periods = period_cols[:BACKLOG_TABLE_PERIODS]

    # Keep top jobs by total backlog
    dfx = dfx.sort_values("Total_R18", ascending=False).head(BACKLOG_TABLE_TOP_N)

    # Rename JobSearch -> Job if present
    dfx = rename_cols_if_exist(dfx, {"JobSearch": "Job"})

    keep = [c if c != "JobSearch" else "Job" for c in meta_cols]  # in case rename happened
    keep = [c for c in keep if c in dfx.columns]

    keep += ["Total_R18"] + [c for c in show_periods if c in dfx.columns]
    return dfx[keep]

# -----------------------------
# Context object
# -----------------------------
@dataclass
class ReportContext:
    recipient_email: str
    recipient_name: str = ""
    audience_type: str = ""
    as_of_str: str = ""          # "2026-02-01"
    snapshot_str: str = ""       # "2026-01-01"
    fiscal_year: int = 0
    div_key: str = "*"
    mkt_key: str = "*"


# -----------------------------
# Public API
# -----------------------------
def build_report_pdf(
    *,
    dfs: List[pd.DataFrame],
    sheet_names: List[str],
    ctx: ReportContext,
    out_dir: str | Path,
    template_path: str | Path | None = None,
    keep_html: bool = True,
) -> Path:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    charts_dir = out_dir / "charts"
    charts_dir.mkdir(parents=True, exist_ok=True)

    data = _dfs_to_named_dict(dfs, sheet_names)
    view = make_view(data)

    insights = compute_insights(view)
    chart_paths = make_charts(view, insights, charts_dir)

    html_path = out_dir / f"Monthly Pulse - {ctx.recipient_email} - Report.html"
    render_html(
        html_path=html_path,
        data=view,  # <-- canonical keys
        ctx=ctx,
        insights=insights,
        chart_paths=chart_paths,
        template_path=template_path,
    )

    pdf_path = out_dir / f"Monthly Pulse - {ctx.recipient_email} - Report.pdf"
    html_to_pdf(html_path, pdf_path)

    if not keep_html:
        try:
            html_path.unlink(missing_ok=True)
        except Exception:
            pass

    return pdf_path

def make_view(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    return {
        "backlog":  data.get("Rolling 18 Backlog", pd.DataFrame()),
        "wip":      data.get("WIP", pd.DataFrame()),
        "rev_month":data.get("Job Revenue (Month)", pd.DataFrame()),
        "rev_ytd":  data.get("Job Revenue (YTD)", pd.DataFrame()),
        "is_month": data.get("Income Statement (Month)", pd.DataFrame()),
        "is_ytd":   data.get("Income Statement (YTD)", pd.DataFrame()),
    }

CANON_TO_SHEET = {
    "rev_month": "Job Revenue (Month)",
    "rev_ytd":   "Job Revenue (YTD)",
}

def get_required_col(df: pd.DataFrame, sheet_name: str, col_name: str) -> str:
    """Return exact or case-insensitive match; raise if missing."""
    if col_name in df.columns:
        return col_name

    lower = {str(c).lower(): c for c in df.columns}
    if col_name.lower() in lower:
        return lower[col_name.lower()]

    raise KeyError(f"[{sheet_name}] Required column not found: {col_name}. Available: {list(df.columns)}")


def get_money_col_for_view_key(view_key: str, df: pd.DataFrame) -> str:
    sheet = CANON_TO_SHEET[view_key]
    col = MONEY_COL_BY_SHEET[sheet]
    return get_required_col(df, sheet, col)


def get_job_col_for_view_key(view_key: str, df: pd.DataFrame) -> str:
    sheet = CANON_TO_SHEET[view_key]
    col = JOB_COL_BY_SHEET[sheet]
    return get_required_col(df, sheet, col)

def df_to_html(df: pd.DataFrame, max_rows: int = 20, table_key: str | None = None) -> str:
    dfx = df.head(max_rows).copy()

    def is_percent_col(colname: str) -> bool:
        c = colname.lower()
        return any(k in c for k in ["%", "percent", "pct"]) or c.endswith("_pct") or c.endswith("_percent")

    def is_money_col(colname: str) -> bool:
        c = colname.lower()
        money_keywords = [
            "revenue", "earnedrevenue", "cost", "amount", "amt",
            "budget", "actual", "variance", "backlog", "contract",
            "billed", "billing", "billings", "paid",
            "profit", "businessplan", "business_plan", "plan"
        ]
        return any(k in c for k in money_keywords)

    def format_money(v):
        return "" if pd.isna(v) else f"${v:,.0f}"

    def format_number(v):
        return "" if pd.isna(v) else f"{v:,.2f}"

    def format_percent(v):
        if pd.isna(v):
            return ""
        return f"{v:.1%}" if abs(v) <= 1 else f"{v:,.2f}"

    for c in dfx.columns:
        s = pd.to_numeric(dfx[c], errors="coerce")
        if s.notna().sum() == 0:
            continue

        cname = str(c)

        # Special rule: backlog table period columns should ALWAYS be currency
        # Because their headers are dates like '25-Feb' and won't match money keywords.
        if table_key == "backlog":
            # Keep meta cols as-is; treat everything else as currency
            if cname not in ("Job", "JobSearch", "Division", "MarketSegment"):
                dfx[c] = s.map(format_money)
            continue

        # Special rule: income statement "BusinessPlan" should always be currency
        if table_key in ("is_month", "is_ytd"):
            if cname.lower() in ("businessplan", "business_plan", "plan", "budget"):
                dfx[c] = s.map(format_money)
                continue

        # WIP: billings_to_complete should be currency
        if table_key == "wip" and cname.lower() == "billings_to_complete":
            dfx[c] = s.map(format_money)
            continue

        if is_percent_col(cname):
            dfx[c] = s.map(format_percent)
        elif is_money_col(cname):
            dfx[c] = s.map(format_money)
        else:
            dfx[c] = s.map(format_number)

    return dfx.to_html(index=False, border=0, classes="df")


def _dfs_to_named_dict(dfs: List[pd.DataFrame], sheet_names: List[str]) -> Dict[str, pd.DataFrame]:
    out: Dict[str, pd.DataFrame] = {}
    for i, name in enumerate(sheet_names):
        out[name] = dfs[i] if i < len(dfs) else pd.DataFrame()
    return out


# -----------------------------
# Insights (lightweight starter)
# -----------------------------
def compute_insights(data: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    insights: Dict[str, Any] = {}

    backlog = data.get("backlog", pd.DataFrame())
    if not backlog.empty and backlog.shape[1] > 3:
        numeric = backlog.iloc[:, 3:].apply(pd.to_numeric, errors="coerce").fillna(0)
        insights["backlog_total_all_periods"] = float(numeric.values.sum())
        insights["backlog_by_period"] = numeric.sum(axis=0)

    wip = data.get("wip", pd.DataFrame())
    if not wip.empty:
        insights["wip_job_count"] = int(len(wip))

    # Revenue (Month/YTD) totals (try to guess amount column)
    for view_key, out_key in [
        ("rev_month", "rev_month_total"),
        ("rev_ytd", "rev_ytd_total"),
    ]:
        df = data.get(view_key, pd.DataFrame())
        if df.empty:
            continue

        amt_col = get_money_col_for_view_key(view_key, df)
        insights[out_key] = float(pd.to_numeric(df[amt_col], errors="coerce").fillna(0).sum())

    return insights


def _find_money_col(df: pd.DataFrame) -> Optional[str]:
    # broad but safe; you can tighten this once you confirm headers
    candidates = ["Revenue", "ActualRevenue", "Amount", "Total", "Value", "RevenueAmount"]
    lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower:
            return lower[cand.lower()]
    # fallback: pick the first numeric column that is not an ID-like column
    for c in df.columns:
        if any(k in str(c).lower() for k in ["job", "contract", "project", "name", "desc"]):
            continue
        s = pd.to_numeric(df[c], errors="coerce")
        if s.notna().sum() > 0:
            return c
    return None


# -----------------------------
# Charts
# -----------------------------
def make_charts(data: Dict[str, pd.DataFrame], insights: Dict[str, Any], out_dir: Path) -> Dict[str, str]:
    chart_paths: Dict[str, str] = {}

    if _HAS_SEABORN:
        sns.set_theme()

    # Backlog by period
    if "backlog_by_period" in insights:
        s = insights["backlog_by_period"]
        # Remove aggregate year columns and take only Rolling 18
        s = s.copy()
        s = s[~pd.Index(s.index.astype(str)).map(is_year_label)]
        s = s.iloc[:BACKLOG_CHART_PERIODS]

        fig = plt.figure(figsize=(10, 3.2))
        ax = plt.gca()
        ax.bar(range(len(s.index)), s.values)
        ax.set_title("Backlog by Period")
        ax.set_xticks(range(len(s.index)))
        ax.set_xticklabels([str(x) for x in s.index], rotation=45, ha="right")
        fig.tight_layout()

        p = out_dir / "backlog_by_period.png"
        fig.savefig(p, dpi=200)
        plt.close(fig)
        chart_paths["backlog_by_period"] = str(p)

    # Top 10 jobs by revenue (Month and YTD) if possible
    for sheet_key, tag in [("rev_month", "month"), ("rev_ytd", "ytd")]:
        df = data.get(sheet_key, pd.DataFrame())
        if df.empty:
            continue

        # Use hardcoded schema for revenue sheets
        job_col = get_job_col_for_view_key(sheet_key, df)
        amt_col = get_money_col_for_view_key(sheet_key, df)

        if not job_col or not amt_col:
            continue

        tmp = df[[job_col, amt_col]].copy()
        tmp[amt_col] = pd.to_numeric(tmp[amt_col], errors="coerce").fillna(0)
        top = tmp.groupby(job_col, as_index=False)[amt_col].sum().sort_values(amt_col, ascending=False).head(10)

        fig = plt.figure(figsize=(10, 4))
        ax = plt.gca()
        ax.barh(top[job_col].astype(str)[::-1], top[amt_col][::-1])
        ax.set_title(f"Top 10 Jobs by Revenue ({'Current Month' if tag=='month' else 'YTD'})")
        fig.tight_layout()

        p = out_dir / f"top10_rev_{tag}.png"
        fig.savefig(p, dpi=200)
        plt.close(fig)
        chart_paths[f"top10_rev_{tag}"] = str(p)

    return chart_paths


def _find_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    lower = {str(c).lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower:
            return lower[cand.lower()]
    for cand in candidates:
        for c in df.columns:
            if cand.lower() in str(c).lower():
                return c
    return None


# -----------------------------
# HTML rendering
# -----------------------------
def render_html(
    *,
    html_path: Path,
    data: Dict[str, pd.DataFrame],
    ctx: ReportContext,
    insights: Dict[str, Any],
    chart_paths: Dict[str, str],
    template_path: str | Path | None,
) -> None:
    html_path.parent.mkdir(parents=True, exist_ok=True)

    if template_path is None:
        template_path = Path(__file__).resolve().parent / "report_template.html"
    else:
        template_path = Path(template_path)

    env = Environment(
        loader=FileSystemLoader(str(template_path.parent)),
        autoescape=select_autoescape(["html", "xml"]),
    )
    tpl = env.get_template(template_path.name)

    # Make chart paths relative to the HTML file folder (portable)
    safe_charts: Dict[str, str] = {}
    for k, p in chart_paths.items():
        rel = os.path.relpath(p, start=str(html_path.parent))
        safe_charts[k] = Path(rel).as_posix()   # HTML-friendly slashes

    tables: Dict[str, str] = {}
    for name, df in data.items():
        df_for_display = df

        # Apply per-table presentation rules based on canonical keys
        if name in ("rev_month", "rev_ytd"):
            df_for_display = prepare_rev_preview(df_for_display)

        if name == "wip":
            df_for_display = prepare_wip_preview(df_for_display)

        if name == "backlog":
            df_for_display = prepare_backlog_table(df_for_display)

        if df_for_display is None or df_for_display.empty:
            tables[name] = "<em>No data.</em>"
        else:
            # if you already switched to df_to_html(...) for currency formatting, use it here
            tables[name] = df_to_html(df_for_display, max_rows=20, table_key=name)

    rendered = tpl.render(ctx=ctx, insights=insights, charts=safe_charts, tables=tables)
    html_path.write_text(rendered, encoding="utf-8")

# -----------------------------
# HTML -> PDF
# -----------------------------
def html_to_pdf(html_path: Path, pdf_path: Path) -> None:
    html_path = html_path.resolve()
    pdf_path = pdf_path.resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(html_path.as_uri(), wait_until="networkidle")
        page.pdf(
            path=str(pdf_path),
            format="Letter",
            print_background=True,
            margin={"top": "0.6in", "right": "0.6in", "bottom": "0.6in", "left": "0.6in"},
        )
        browser.close()


# -----------------------------
# CLI (testing)
# -----------------------------
def main():
    import argparse
    parser = argparse.ArgumentParser()

    parser.add_argument("--in_dir", required=True, help="Folder containing backlog.parquet, wip.parquet, rev_month.parquet, rev_ytd.parquet, is_month.parquet, is_ytd.parquet")
    parser.add_argument("--out_dir", required=True)

    parser.add_argument("--email", default="dev@test.com")
    parser.add_argument("--as_of", default="2026-02-01")
    parser.add_argument("--snapshot", default="2026-01-01")
    parser.add_argument("--fy", type=int, default=2026)
    parser.add_argument("--div", default="*")
    parser.add_argument("--mkt", default="*")

    args = parser.parse_args()

    in_dir = Path(args.in_dir)

    dfs = [
        pd.read_parquet(in_dir / "backlog.parquet"),
        pd.read_parquet(in_dir / "wip.parquet"),
        pd.read_parquet(in_dir / "rev_month.parquet"),
        pd.read_parquet(in_dir / "rev_ytd.parquet"),
        pd.read_parquet(in_dir / "is_month.parquet"),
        pd.read_parquet(in_dir / "is_ytd.parquet"),
    ]

    SHEETS = [
        "Rolling 18 Backlog",
        "WIP",
        "Job Revenue (Month)",
        "Job Revenue (YTD)",
        "Income Statement (Month)",
        "Income Statement (YTD)",
    ]

    ctx = ReportContext(
        recipient_email=args.email,
        as_of_str=args.as_of,
        snapshot_str=args.snapshot,
        fiscal_year=args.fy,
        div_key=args.div,
        mkt_key=args.mkt,
    )

    pdf = build_report_pdf(dfs=dfs, sheet_names=SHEETS, ctx=ctx, out_dir=args.out_dir, keep_html=True)
    print(f"Created PDF: {pdf}")



if __name__ == "__main__":
    main()
