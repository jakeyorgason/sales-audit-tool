import os
from typing import Any, Optional

import pandas as pd
import requests
import streamlit as st

from sales_audit_ingestion import SalesAuditEngine


st.set_page_config(
    page_title="Free Amazon Ads Audit | Evolved Commerce",
    page_icon="🟧",
    layout="wide",
)

# =========================================================
# SESSION STATE
# =========================================================
DEFAULT_STATE = {
    "sales_audit_results": {},
    "created_report": None,
    "unlock_complete": False,
    "lead_name": "",
    "lead_email": "",
    "lead_phone": "",
    "lead_brand_name": "",
}

for _k, _v in DEFAULT_STATE.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# =========================================================
# HELPERS
# =========================================================
def safe_dict(value: Any) -> dict:
    return value if isinstance(value, dict) else {}


def safe_df(value: Any) -> pd.DataFrame:
    return value if isinstance(value, pd.DataFrame) else pd.DataFrame()


def get_number(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def load_logo_path() -> Optional[str]:
    possible_paths = [
        "assets/ec_logo.png",
        "assets/ec_logo.jpg",
        "assets/ec_logo.jpeg",
        "assets/logo.png",
        "assets/logo.jpg",
        "ec_logo.png",
        "logo.png",
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None


def format_currency(value: float) -> str:
    return f"${get_number(value):,.2f}"


def format_percent(value: float) -> str:
    return f"{get_number(value):,.2f}%"


def format_number(value: float) -> str:
    return f"{get_number(value):,.2f}"


def render_metric_card(label: str, value: str, tone: str = "brand", small: bool = False) -> None:
    value_class = "metric-value small" if small else "metric-value"
    st.markdown(
        f'''
        <div class="metric-card {tone}">
            <div class="metric-label">{label}</div>
            <div class="{value_class}">{value}</div>
        </div>
        ''',
        unsafe_allow_html=True,
    )


def tone_from_health(status: str) -> str:
    status = str(status).strip().lower()
    if status == "healthy":
        return "good"
    if status == "mixed":
        return "warn"
    return "bad"


def simplify_term_table(df: pd.DataFrame, term_col: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    if "acos" in out.columns and "acos_pct" not in out.columns:
        out["acos_pct"] = out["acos"] * 100

    keep_cols = [c for c in [term_col, "spend", "sales", "acos_pct"] if c in out.columns]
    out = out[keep_cols].copy()

    out = out.rename(columns={
        term_col: "term",
        "spend": "spend",
        "sales": "sales",
        "acos_pct": "acos",
    })

    out["spend"] = pd.to_numeric(out["spend"], errors="coerce").fillna(0)
    out["sales"] = pd.to_numeric(out["sales"], errors="coerce").fillna(0)
    out["acos"] = pd.to_numeric(out["acos"], errors="coerce").fillna(0)

    out["spend"] = out["spend"].map(lambda x: f"${x:,.2f}")
    out["sales"] = out["sales"].map(lambda x: f"${x:,.2f}")
    out["acos"] = out["acos"].map(lambda x: f"{x:,.2f}%")

    return out.reset_index(drop=True)


def simplify_campaign_table(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()
    keep = [c for c in ["campaign_name", "spend", "sales", "acos_pct", "campaign_status"] if c in out.columns]
    out = out[keep].copy()

    if "spend" in out.columns:
        out["spend"] = pd.to_numeric(out["spend"], errors="coerce").fillna(0).round(2)
    if "sales" in out.columns:
        out["sales"] = pd.to_numeric(out["sales"], errors="coerce").fillna(0).round(2)
    if "acos_pct" in out.columns:
        out["acos_pct"] = pd.to_numeric(out["acos_pct"], errors="coerce").fillna(0).round(2)

    return out.sort_values(["spend", "sales"], ascending=[False, False]).reset_index(drop=True)


def build_sheet_term_table(df: pd.DataFrame, term_col: str) -> pd.DataFrame:
    '''
    Build raw numeric rows for the Google Sheet payload.

    NOTE:
    For this audit workflow, the term-level spend/sales tables are arriving in cents,
    while the KPI summary is already in dollars. We normalize to dollars here.
    '''
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    if "acos" in out.columns and "acos_pct" not in out.columns:
        out["acos_pct"] = out["acos"] * 100

    keep_cols = [c for c in [term_col, "spend", "sales", "acos_pct"] if c in out.columns]
    out = out[keep_cols].copy()

    out = out.rename(columns={
        term_col: "term",
        "spend": "spend",
        "sales": "sales",
        "acos_pct": "acos",
    })

    out["spend"] = pd.to_numeric(out["spend"], errors="coerce").fillna(0)
    out["sales"] = pd.to_numeric(out["sales"], errors="coerce").fillna(0)
    out["acos"] = pd.to_numeric(out["acos"], errors="coerce").fillna(0)

    out["spend"] = (out["spend"] / 100.0).round(2)
    out["sales"] = (out["sales"] / 100.0).round(2)
    out["acos"] = out["acos"].round(2)

    return out.reset_index(drop=True)


def normalize_records_for_sheet(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []

    out = df.copy()
    out = out.replace({pd.NA: None})
    out = out.where(pd.notnull(out), None)

    def parse_numeric_like(value):
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return value

        text = str(value).strip()
        if text == "":
            return None

        negative = text.startswith("(") and text.endswith(")")
        text = text.replace("(", "").replace(")", "")
        text = text.replace("$", "").replace(",", "").replace("%", "").strip()

        try:
            num = float(text)
            return -num if negative else num
        except ValueError:
            return value

    for col in out.columns:
        col_l = str(col).lower().strip()
        out[col] = out[col].map(parse_numeric_like)

        if col_l in {"ctr", "cvr", "acos", "acos_pct"}:
            out[col] = (
                pd.to_numeric(out[col], errors="coerce")
                .fillna(0)
                .map(lambda x: round(float(x) / 100.0, 6))
            )
        elif col_l in {"roas"}:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).map(lambda x: round(float(x), 2))
        elif col_l in {"cpc"}:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).map(lambda x: round(float(x), 2))
        elif col_l in {"spend", "sales", "ad_sales", "total_sales", "organic_sales", "ntb_sales"}:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).map(lambda x: round(float(x), 2))
        elif col_l in {"impressions", "clicks", "orders", "units_ordered", "sessions", "ntb_orders"}:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).map(lambda x: int(round(float(x))))
        elif "pct" in col_l or "percent" in col_l:
            out[col] = (
                pd.to_numeric(out[col], errors="coerce")
                .fillna(0)
                .map(lambda x: round(float(x) / 100.0, 6))
            )
        elif pd.api.types.is_numeric_dtype(out[col]):
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    return out.to_dict("records")


def create_google_sheet_report(
    brand_name: str,
    report_name: str,
    date_range_label: str,
    kpi_summary: dict,
    waste_summary: dict,
    match_type_revenue_rows: list[dict],
    match_type_inefficient_rows: list[dict],
    campaign_rows: list[dict],
    campaign_type_rows: list[dict],
    top_keyword_rows: list[dict],
    top_search_term_rows: list[dict],
    waste_keyword_rows: list[dict],
    waste_search_term_rows: list[dict],
    winner_keyword_rows: list[dict],
    winner_search_term_rows: list[dict],
    targeting_data_rows: list[dict],
    search_term_data_rows: list[dict],
) -> dict:
    webhook_url = os.getenv("GOOGLE_SHEET_WEBHOOK_URL")
    template_id = os.getenv("GOOGLE_SHEETS_TEMPLATE_ID")
    destination_folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

    missing = []
    if not webhook_url:
        missing.append("GOOGLE_SHEET_WEBHOOK_URL")
    if not template_id:
        missing.append("GOOGLE_SHEETS_TEMPLATE_ID")
    if not destination_folder_id:
        missing.append("GOOGLE_DRIVE_FOLDER_ID")

    if missing:
        raise RuntimeError(f"Missing environment variables: {', '.join(missing)}")

    payload = {
        "templateId": template_id,
        "destinationFolderId": destination_folder_id,
        "reportName": report_name,
        "brandName": brand_name,
        "dateRangeLabel": date_range_label,
        "kpiSummary": kpi_summary,
        "wasteSummary": waste_summary,
        "matchTypeRevenueRows": match_type_revenue_rows,
        "matchTypeInefficientRows": match_type_inefficient_rows,
        "campaignRows": campaign_rows,
        "campaignTypeRows": campaign_type_rows,
        "topKeywordRows": top_keyword_rows,
        "topSearchTermRows": top_search_term_rows,
        "wasteKeywordRows": waste_keyword_rows,
        "wasteSearchTermRows": waste_search_term_rows,
        "winnerKeywordRows": winner_keyword_rows,
        "winnerSearchTermRows": winner_search_term_rows,
        "targetingDataRows": targeting_data_rows,
        "searchTermDataRows": search_term_data_rows,
    }

    response = requests.post(webhook_url, json=payload, timeout=180)
    response.raise_for_status()

    data = response.json()
    if not data.get("success"):
        raise RuntimeError(data.get("error", "Unknown Apps Script error"))

    return data


# =========================================================
# STYLING
# =========================================================
        st.markdown(
            '''
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
        
                :root {
                    --ec-orange: #f47322;
                    --ec-orange-dark: #d85f16;
                    --ec-orange-soft: #fff3ea;
                    --ec-black: #111111;
                    --ec-charcoal: #252525;
                    --ec-muted: #686868;
                    --ec-border: #e9e2da;
                    --ec-bg: #fbfaf7;
                    --ec-card: #ffffff;
                    --ec-green: #16a34a;
                    --ec-yellow: #d97706;
                    --ec-red: #dc2626;
                }
        
                html, body, [class*="css"] {
                    font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
                }
        
                .stApp {
                    background:
                        radial-gradient(circle at top left, rgba(244, 115, 34, 0.12), transparent 30rem),
                        linear-gradient(180deg, #ffffff 0%, var(--ec-bg) 42%, #ffffff 100%);
                    color: var(--ec-black);
                }
        
                .main > div {
                    padding-top: 1rem;
                }
        
                .block-container {
                    padding-top: 1rem;
                    padding-bottom: 2.25rem;
                    max-width: 1360px;
                }
        
                /* Sidebar */
                [data-testid="stSidebar"] {
                    background: #111111;
                    border-right: 1px solid rgba(255,255,255,0.08);
                }
        
                [data-testid="stSidebar"] * {
                    color: rgba(255,255,255,0.88);
                }
        
                [data-testid="stSidebar"] h1,
                [data-testid="stSidebar"] h2,
                [data-testid="stSidebar"] h3,
                [data-testid="stSidebar"] strong {
                    color: #ffffff !important;
                }
        
                [data-testid="stSidebar"] hr {
                    border-color: rgba(255,255,255,0.14);
                }
        
                [data-testid="stSidebar"] .stButton > button {
                    background: var(--ec-orange) !important;
                    border-color: var(--ec-orange) !important;
                    color: #ffffff !important;
                    box-shadow: 0 14px 28px rgba(244, 115, 34, 0.22);
                }
        
                [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
                [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] li {
                    font-size: 0.91rem;
                    line-height: 1.5;
                }
        
                /* Header / landing hero */
                .hero-wrap {
                    background: #ffffff;
                    border: 1px solid var(--ec-border);
                    border-radius: 30px;
                    padding: 44px 46px 40px 46px;
                    margin-bottom: 1.2rem;
                    box-shadow: 0 26px 80px rgba(17, 17, 17, 0.08);
                    position: relative;
                    overflow: hidden;
                }
        
                .hero-wrap:before {
                    content: "";
                    position: absolute;
                    inset: 0;
                    background:
                        radial-gradient(circle at top right, rgba(244, 115, 34, 0.14), transparent 23rem),
                        linear-gradient(90deg, rgba(255,255,255,0.00), rgba(244,115,34,0.05));
                    pointer-events: none;
                }
        
                .hero-wrap:after {
                    content: "";
                    position: absolute;
                    top: 0;
                    left: 0;
                    right: 0;
                    height: 7px;
                    background: linear-gradient(90deg, var(--ec-orange), #ffb36f);
                }
        
                .hero-inner {
                    position: relative;
                    z-index: 1;
                }
        
                .hero-kicker {
                    display: inline-flex;
                    align-items: center;
                    gap: 8px;
                    padding: 7px 13px;
                    background: var(--ec-orange-soft);
                    border: 1px solid rgba(244, 115, 34, 0.25);
                    border-radius: 999px;
                    color: var(--ec-orange-dark);
                    font-size: 0.76rem;
                    font-weight: 900;
                    letter-spacing: 0.075em;
                    text-transform: uppercase;
                    margin-bottom: 15px;
                }
        
                .brand-title {
                    font-size: clamp(2.45rem, 5vw, 5.25rem);
                    font-weight: 900;
                    letter-spacing: -0.075em;
                    line-height: 0.91;
                    margin: 0;
                    color: var(--ec-black);
                    max-width: 980px;
                }
        
                .brand-title .accent {
                    color: var(--ec-orange);
                    display: inline-block;
                }
        
                .brand-subtitle {
                    font-size: 1.08rem;
                    color: var(--ec-muted);
                    margin-top: 1.05rem;
                    max-width: 870px;
                    line-height: 1.65;
                    font-weight: 500;
                }
        
                .hero-pill-row {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 9px;
                    margin-top: 20px;
                }
        
                .hero-pill {
                    background: #111111;
                    color: #ffffff;
                    border-radius: 999px;
                    padding: 8px 12px;
                    font-size: 0.78rem;
                    font-weight: 800;
                    letter-spacing: 0.01em;
                }
        
                .logo-card {
                    background: #ffffff;
                    border: 1px solid var(--ec-border);
                    border-radius: 26px;
                    padding: 24px;
                    box-shadow: 0 18px 50px rgba(17,17,17,0.065);
                    min-height: 210px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }
        
                .trust-strip {
                    display: grid;
                    grid-template-columns: repeat(3, 1fr);
                    gap: 14px;
                    margin: 0.4rem 0 1.2rem 0;
                }
        
                .trust-card {
                    background: rgba(255,255,255,0.82);
                    border: 1px solid var(--ec-border);
                    border-radius: 22px;
                    padding: 17px 18px;
                    box-shadow: 0 14px 36px rgba(17,17,17,0.045);
                }
        
                .trust-card strong {
                    display: block;
                    color: var(--ec-black);
                    font-size: 0.96rem;
                    font-weight: 900;
                    letter-spacing: -0.02em;
                    margin-bottom: 4px;
                }
        
                .trust-card span {
                    display: block;
                    color: var(--ec-muted);
                    font-size: 0.9rem;
                    line-height: 1.45;
                    font-weight: 500;
                }
        
                /* Section headings */
                .section-title {
                    font-size: 1.32rem;
                    font-weight: 900;
                    letter-spacing: -0.035em;
                    color: var(--ec-black);
                    margin-bottom: 0.25rem;
                }
        
                .section-title:before {
                    content: "";
                    display: inline-block;
                    width: 9px;
                    height: 9px;
                    background: var(--ec-orange);
                    border-radius: 999px;
                    margin-right: 9px;
                    transform: translateY(-1px);
                }
        
                .section-note {
                    font-size: 0.96rem;
                    color: var(--ec-muted);
                    margin-bottom: 1rem;
                    line-height: 1.5;
                    font-weight: 500;
                }
        
                /* Metric cards */
                .metric-card {
                    background: var(--ec-card);
                    border: 1px solid var(--ec-border);
                    border-radius: 22px;
                    padding: 17px 18px;
                    box-shadow: 0 14px 38px rgba(17,17,17,0.055);
                    min-height: 104px;
                    transition: transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease;
                }
        
                .metric-card:hover {
                    transform: translateY(-2px);
                    box-shadow: 0 20px 48px rgba(17,17,17,0.08);
                    border-color: rgba(244, 115, 34, 0.38);
                }
        
                .metric-label {
                    font-size: 0.72rem;
                    color: var(--ec-muted);
                    margin-bottom: 0.38rem;
                    font-weight: 900;
                    text-transform: uppercase;
                    letter-spacing: 0.055em;
                }
        
                .metric-value {
                    font-size: 1.68rem;
                    line-height: 1.05;
                    font-weight: 900;
                    letter-spacing: -0.045em;
                    color: var(--ec-black);
                    word-break: break-word;
                }
        
                .metric-value.small {
                    font-size: 1.14rem;
                    letter-spacing: -0.025em;
                }
        
                .metric-card.good { border-left: 6px solid var(--ec-green); }
                .metric-card.warn { border-left: 6px solid var(--ec-yellow); }
                .metric-card.bad { border-left: 6px solid var(--ec-red); }
                .metric-card.brand { border-left: 6px solid var(--ec-orange); }
        
                /* Health + summary */
                .status-pill {
                    display: inline-flex;
                    align-items: center;
                    padding: 9px 14px;
                    border-radius: 999px;
                    font-size: 0.82rem;
                    font-weight: 900;
                    letter-spacing: 0.035em;
                    text-transform: uppercase;
                    margin-bottom: 0.75rem;
                }
        
                .status-pill.good {
                    background: #ecfdf3;
                    color: #047857;
                    border: 1px solid #bbf7d0;
                }
        
                .status-pill.warn {
                    background: #fff7ed;
                    color: #b45309;
                    border: 1px solid #fed7aa;
                }
        
                .status-pill.bad {
                    background: #fef2f2;
                    color: #b91c1c;
                    border: 1px solid #fecaca;
                }
        
                .summary-box {
                    background: #ffffff;
                    border: 1px solid var(--ec-border);
                    border-radius: 22px;
                    padding: 20px 22px;
                    margin-bottom: 16px;
                    box-shadow: 0 14px 38px rgba(17,17,17,0.05);
                    color: var(--ec-charcoal);
                    line-height: 1.62;
                }
        
                .summary-box strong {
                    color: var(--ec-black);
                    font-weight: 900;
                }
        
                .upload-note {
                    color: var(--ec-muted);
                    font-size: 0.93rem;
                    margin-top: 0.1rem;
                    margin-bottom: 0.75rem;
                    font-weight: 700;
                }
        
                .lead-card {
                    background:
                        radial-gradient(circle at top right, rgba(244,115,34,0.16), transparent 20rem),
                        #111111;
                    color: #ffffff;
                    border-radius: 30px;
                    padding: 34px 36px;
                    box-shadow: 0 28px 80px rgba(17,17,17,0.18);
                    margin: 1.5rem 0 1.3rem 0;
                    border: 1px solid rgba(255,255,255,0.08);
                }
        
                .lead-card h2 {
                    color: #ffffff;
                    font-size: 2rem;
                    line-height: 1.03;
                    letter-spacing: -0.055em;
                    margin: 0 0 0.65rem 0;
                    font-weight: 900;
                }
        
                .lead-card p {
                    color: rgba(255,255,255,0.78);
                    font-size: 1rem;
                    line-height: 1.55;
                    margin-bottom: 0;
                    font-weight: 500;
                }
        
                .lead-card .mini {
                    display: inline-block;
                    color: #ffffff;
                    background: var(--ec-orange);
                    border-radius: 999px;
                    padding: 7px 12px;
                    font-size: 0.76rem;
                    font-weight: 900;
                    letter-spacing: 0.06em;
                    text-transform: uppercase;
                    margin-bottom: 13px;
                }
        
                /* Native Streamlit widgets */
                [data-testid="stFileUploader"] {
                    background: #ffffff;
                    border: 1px dashed rgba(244,115,34,0.48);
                    border-radius: 20px;
                    padding: 0.85rem;
                    box-shadow: 0 10px 26px rgba(17,17,17,0.035);
                }
        
                .stTextInput input {
                    border-radius: 14px !important;
                    border: 1px solid var(--ec-border) !important;
                    min-height: 44px;
                }
        
                .stTextInput input:focus {
                    border-color: var(--ec-orange) !important;
                    box-shadow: 0 0 0 3px rgba(244,115,34,0.14) !important;
                }
        
                .stButton > button {
                    border-radius: 999px !important;
                    border: 1px solid #111111 !important;
                    background: #111111 !important;
                    color: #ffffff !important;
                    font-weight: 900 !important;
                    padding: 0.78rem 1.25rem !important;
                    box-shadow: 0 14px 30px rgba(17,17,17,0.16);
                    transition: transform 0.15s ease, box-shadow 0.15s ease, background 0.15s ease;
                }
        
                .stButton > button:hover {
                    background: var(--ec-orange) !important;
                    border-color: var(--ec-orange) !important;
                    transform: translateY(-1px);
                    box-shadow: 0 18px 38px rgba(244,115,34,0.24);
                }
        
                .stButton > button[kind="primary"],
                button[kind="primary"] {
                    background: var(--ec-orange) !important;
                    border-color: var(--ec-orange) !important;
                    color: #ffffff !important;
                }
        
                .stButton > button[kind="primary"]:hover,
                button[kind="primary"]:hover {
                    background: var(--ec-orange-dark) !important;
                    border-color: var(--ec-orange-dark) !important;
                }
        
                .stButton > button:disabled {
                    background: #d8d3cc !important;
                    border-color: #d8d3cc !important;
                    color: #ffffff !important;
                    box-shadow: none !important;
                }
        
                div[data-testid="stDataFrame"] {
                    border: 1px solid var(--ec-border);
                    border-radius: 18px;
                    overflow: hidden;
                    box-shadow: 0 12px 34px rgba(17,17,17,0.045);
                }
        
                div[data-testid="stExpander"] {
                    background: rgba(255,255,255,0.85);
                    border: 1px solid var(--ec-border);
                    border-radius: 18px;
                    box-shadow: 0 10px 28px rgba(17,17,17,0.04);
                    overflow: hidden;
                }
        
                div[data-testid="stAlert"] {
                    border-radius: 16px;
                    border: 1px solid var(--ec-border);
                }
        
                hr {
                    border-color: var(--ec-border);
                    margin: 1.55rem 0;
                }
        
                a {
                    color: var(--ec-orange-dark);
                    font-weight: 800;
                }
        
                @media (max-width: 900px) {
                    .trust-strip {
                        grid-template-columns: 1fr;
                    }
        
                    .hero-wrap {
                        padding: 34px 26px;
                    }
        
                    .brand-title {
                        font-size: 2.65rem;
                    }
        
                    .lead-card {
                        padding: 28px 24px;
                    }
                }
            </style>
            ''',
            unsafe_allow_html=True,
        )


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown("## Free Audit Checklist")
    st.markdown(
        '''
Upload these reports to generate your audit:

- Bulk Sheet
- Sponsored Products Search Term Report
- Sponsored Products Targeting Report
- Sponsored Products Impression Share Report
- Sales & Traffic Business Report
- Sponsored Brands Campaign Report *(optional)*
'''
    )

    st.markdown("---")

    st.markdown("## What You’ll See")
    st.markdown(
        '''
- Account health verdict
- Executive KPI snapshot
- Wasted spend indicators
- Top keywords and search terms
- Campaign-level performance summary
- Personalized Google Sheets audit
'''
    )

    st.markdown("---")

    st.markdown("## Best Results")
    st.markdown(
        '''
Use matching date ranges where possible. The Sponsored Brands report helps populate new-to-brand context, but it is not required.
'''
    )

    if st.button("Start over", use_container_width=True):
        st.session_state["sales_audit_results"] = {}
        st.session_state["created_report"] = None
        st.session_state["unlock_complete"] = False
        st.session_state["lead_name"] = ""
        st.session_state["lead_email"] = ""
        st.session_state["lead_phone"] = ""
        st.session_state["lead_brand_name"] = ""
        st.rerun()


# =========================================================
# HEADER
# =========================================================
logo_path = load_logo_path()
header_left, header_right = st.columns([1.15, 5.85], gap="large")

with header_left:
    if logo_path:
        st.markdown('<div class="logo-card">', unsafe_allow_html=True)
        st.image(logo_path, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

with header_right:
    st.markdown(
        '''
        <div class="hero-wrap">
            <div class="hero-inner">
                <div class="hero-kicker">Free Amazon Advertising Audit</div>
                <div class="brand-title">
                    Find wasted spend.<br>
                    <span class="accent">Unlock growth.</span>
                </div>
                <div class="brand-subtitle">
                    Upload your recent Amazon reports and get a personalized audit that highlights
                    inefficient spend, strongest revenue drivers, top search opportunities, and
                    campaign-level issues your team can act on.
                </div>
                <div class="hero-pill-row">
                    <span class="hero-pill">Amazon Ads</span>
                    <span class="hero-pill">Wasted Spend</span>
                    <span class="hero-pill">Search Terms</span>
                    <span class="hero-pill">Campaign Health</span>
                    <span class="hero-pill">Free Report</span>
                </div>
            </div>
        </div>
        ''',
        unsafe_allow_html=True,
    )

st.markdown(
    '''
    <div class="trust-strip">
        <div class="trust-card">
            <strong>Built for Amazon brands</strong>
            <span>Review ads performance, search terms, targeting, and account efficiency in one guided workflow.</span>
        </div>
        <div class="trust-card">
            <strong>Focused on profitable growth</strong>
            <span>Spot wasted spend, identify winners, and understand where advertising is helping or hurting sales velocity.</span>
        </div>
        <div class="trust-card">
            <strong>Personalized report output</strong>
            <span>Submit your details after the preview to generate a shareable Google Sheets audit for your brand.</span>
        </div>
    </div>
    ''',
    unsafe_allow_html=True,
)


# =========================================================
# INPUTS
# =========================================================
st.markdown('<div class="section-title">Start Your Free Audit</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-note">Enter your brand name and upload the five required Amazon reports. Optional Sponsored Brands data can make the audit more complete.</div>',
    unsafe_allow_html=True,
)

brand_name_input = st.text_input(
    "Brand Name",
    value=st.session_state.get("lead_brand_name", ""),
    placeholder="Your brand name",
    key="sales_audit_brand_name",
)

u1, u2 = st.columns(2)
u3, u4 = st.columns(2)
u5, u6 = st.columns(2)

with u1:
    bulk_file = st.file_uploader("Bulk Sheet", type=["xlsx", "xls", "csv"], key="sales_audit_bulk_file")
with u2:
    impression_share_file = st.file_uploader("SP Impression Share Report", type=["csv", "xlsx", "xls"], key="sales_audit_impression_share_file")
with u3:
    targeting_file = st.file_uploader("SP Targeting Report", type=["csv", "xlsx", "xls"], key="sales_audit_targeting_file")
with u4:
    search_term_file = st.file_uploader("SP Search Term Report", type=["csv", "xlsx", "xls"], key="sales_audit_search_term_file")
with u5:
    business_report_file = st.file_uploader("Sales & Traffic Business Report", type=["csv", "xlsx", "xls"], key="sales_audit_business_report_file")
with u6:
    sb_campaign_file = st.file_uploader("Sponsored Brands Campaign Report (optional)", type=["csv", "xlsx", "xls"], key="sales_audit_sb_campaign_file")

required_ready = all([
    brand_name_input.strip(),
    bulk_file is not None,
    impression_share_file is not None,
    targeting_file is not None,
    search_term_file is not None,
    business_report_file is not None,
])

st.markdown(
    f'<div class="upload-note">{"Ready to generate your audit." if required_ready else "Add your brand name and all five required reports to unlock the audit button."}</div>',
    unsafe_allow_html=True,
)

run_clicked = st.button(
    "Get My Free Audit",
    type="primary",
    use_container_width=True,
    disabled=not required_ready,
)


# =========================================================
# PROCESS AUDIT
# =========================================================
if run_clicked:
    st.session_state["sales_audit_results"] = {}
    st.session_state["created_report"] = None
    st.session_state["unlock_complete"] = False
    st.session_state["lead_brand_name"] = brand_name_input.strip()

    with st.spinner("Running sales audit..."):
        try:
            engine = SalesAuditEngine(
                bulk_file=bulk_file,
                impression_share_file=impression_share_file,
                targeting_file=targeting_file,
                search_term_file=search_term_file,
                business_report_file=business_report_file,
                sb_campaign_file=sb_campaign_file,
                high_acos_threshold=40.0,
                winning_acos_threshold=15.0,
                min_waste_spend=0.01,
                min_winner_orders=1,
            )
            engine.brand_name = brand_name_input.strip()

            results = engine.process()
            results["brand_name"] = brand_name_input.strip()
            st.session_state["sales_audit_results"] = results
        except Exception as exc:
            st.error(f"Sales audit failed: {exc}")

results = safe_dict(st.session_state.get("sales_audit_results", {}))
brand_name = str(results.get("brand_name", st.session_state.get("lead_brand_name", ""))).strip()

if results:
    kpis = safe_dict(results.get("kpi_summary"))
    waste_summary = safe_dict(results.get("waste_summary"))
    health_summary = safe_dict(results.get("health_summary"))
    campaign_summary = safe_df(results.get("campaign_summary"))
    keyword_spend_table = safe_df(results.get("keyword_spend_table"))
    search_term_spend_table = safe_df(results.get("search_term_spend_table"))
    waste_tables = safe_dict(results.get("waste_tables"))
    winner_tables = safe_dict(results.get("winner_tables"))
    narrative = str(results.get("narrative", "")).strip()

    # UI tables
    top_kw = simplify_term_table(keyword_spend_table, "target").head(20)
    top_st = simplify_term_table(search_term_spend_table, "customer_search_term").head(20)
    campaign_view = simplify_campaign_table(campaign_summary).head(20)

    # Raw sheet tables
    top_kw_sheet = build_sheet_term_table(keyword_spend_table, "target").head(20)
    top_st_sheet = build_sheet_term_table(search_term_spend_table, "customer_search_term").head(20)

    waste_kw_sheet = build_sheet_term_table(
        pd.concat(
            [
                safe_df(waste_tables.get("keyword_zero_sale")),
                safe_df(waste_tables.get("keyword_high_acos")),
            ],
            ignore_index=True,
        ).drop_duplicates(),
        "target",
    )

    waste_st_sheet = build_sheet_term_table(
        pd.concat(
            [
                safe_df(waste_tables.get("search_zero_sale")),
                safe_df(waste_tables.get("search_high_acos")),
            ],
            ignore_index=True,
        ).drop_duplicates(),
        "customer_search_term",
    )

    kw_winners_sheet = build_sheet_term_table(
        safe_df(winner_tables.get("keyword_winners")),
        "target",
    ).head(20)

    st_winners_sheet = build_sheet_term_table(
        safe_df(winner_tables.get("search_winners")),
        "customer_search_term",
    ).head(20)

    if brand_name:
        st.markdown(f"### {brand_name}")

    st.markdown('<div class="section-title">Account Health Verdict</div>', unsafe_allow_html=True)
    status = health_summary.get("status", "Unknown")
    tone = tone_from_health(status)
    st.markdown(f'<div class="status-pill {tone}">{status}</div>', unsafe_allow_html=True)
    st.markdown(
        f'''
        <div class="summary-box">
            <strong>Summary:</strong> {health_summary.get("summary", "No summary available.")}
            <br><br>
            <strong>Narrative:</strong> {narrative}
        </div>
        ''',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Executive KPI Snapshot</div>', unsafe_allow_html=True)
    r1 = st.columns(4)
    with r1[0]:
        render_metric_card("Spend", format_currency(kpis.get("spend")), tone="brand")
    with r1[1]:
        render_metric_card("Ad Sales", format_currency(kpis.get("ad_sales")), tone="brand")
    with r1[2]:
        render_metric_card("Total Sales", format_currency(kpis.get("total_sales")), tone="brand")
    with r1[3]:
        render_metric_card("Organic Sales", format_currency(kpis.get("organic_sales")), tone="brand")

    r2 = st.columns(4)
    with r2[0]:
        render_metric_card("ACOS", format_percent(kpis.get("acos_pct")), tone="warn")
    with r2[1]:
        render_metric_card("ROAS", format_number(kpis.get("roas")), tone="good")
    with r2[2]:
        render_metric_card("TACOS", format_percent(kpis.get("tacos_pct")), tone="warn")
    with r2[3]:
        render_metric_card("Wasted Spend", format_currency(waste_summary.get("wasted_spend")), tone="bad")

    # CENTERED LEAD FORM
    st.markdown(
        '''
        <div class="lead-card">
            <span class="mini">Your personalized report is ready to build</span>
            <h2>Get your free Amazon Ads audit report.</h2>
            <p>
                Submit your details and we’ll generate a personalized Google Sheets audit
                with your KPI snapshot, wasted spend review, winning terms, and campaign summary.
            </p>
        </div>
        ''',
        unsafe_allow_html=True,
    )

    if not st.session_state["unlock_complete"]:
        left_pad, center_col, right_pad = st.columns([1, 2, 1])

        with center_col:
            lead_name = st.text_input(
                "Full Name",
                value=st.session_state.get("lead_name", ""),
                key="unlock_name",
            )
            lead_phone = st.text_input(
                "Phone Number",
                value=st.session_state.get("lead_phone", ""),
                key="unlock_phone",
            )
            lead_email = st.text_input(
                "Work Email",
                value=st.session_state.get("lead_email", ""),
                key="unlock_email",
            )
            lead_brand_name = st.text_input(
                "Brand Name",
                value=brand_name or st.session_state.get("lead_brand_name", ""),
                key="unlock_brand",
            )

            unlock_clicked = st.button(
                "Unlock My Personalized Audit",
                use_container_width=True,
                key="unlock_button",
            )

        if unlock_clicked:
            st.session_state["lead_name"] = lead_name
            st.session_state["lead_email"] = lead_email
            st.session_state["lead_phone"] = lead_phone
            st.session_state["lead_brand_name"] = lead_brand_name

            try:
                date_range_label = (results.get("date_range_label") or "").strip() or "MM/DD - MM/DD"

                with st.spinner("Generating your personalized audit..."):
                    created_report = create_google_sheet_report(
                        brand_name=lead_brand_name,
                        report_name=f"{lead_brand_name} - Amazon Ads Audit",
                        date_range_label=date_range_label,
                        kpi_summary=kpis,
                        waste_summary=waste_summary,
                        match_type_revenue_rows=results.get("match_type_revenue_rows", []),
                        match_type_inefficient_rows=results.get("match_type_inefficient_rows", []),
                        campaign_rows=normalize_records_for_sheet(campaign_summary),
                        campaign_type_rows=results.get("campaign_type_rows", []),
                        top_keyword_rows=normalize_records_for_sheet(top_kw_sheet),
                        top_search_term_rows=normalize_records_for_sheet(top_st_sheet),
                        waste_keyword_rows=normalize_records_for_sheet(waste_kw_sheet),
                        waste_search_term_rows=normalize_records_for_sheet(waste_st_sheet),
                        winner_keyword_rows=normalize_records_for_sheet(kw_winners_sheet),
                        winner_search_term_rows=normalize_records_for_sheet(st_winners_sheet),
                        targeting_data_rows=normalize_records_for_sheet(safe_df(results.get("targeting_with_share"))),
                        search_term_data_rows=normalize_records_for_sheet(safe_df(results.get("search_terms"))),
                    )

                st.session_state["created_report"] = created_report
                st.session_state["unlock_complete"] = True
                st.rerun()

            except Exception as exc:
                st.error(f"We hit an issue while creating the audit: {exc}")

    if st.session_state["unlock_complete"] and st.session_state["created_report"]:
        created_report = st.session_state["created_report"]
        st.success("Your personalized audit is ready.")
        st.markdown(
            f'''
            <div class="summary-box">
                <strong>Your Google Sheets audit has been created.</strong><br>
                <a href="{created_report['url']}" target="_blank">Open your personalized Amazon Ads audit →</a>
            </div>
            ''',
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.markdown("### Quick Overview")
    overview_left, overview_right = st.columns(2)

    with overview_left:
        st.markdown("**Top Keywords / Targets**")
        top_kw_brief = top_kw.head(5).copy()
        if not top_kw_brief.empty:
            st.dataframe(top_kw_brief, use_container_width=True, hide_index=True)
        else:
            st.info("No top keyword data available.")

    with overview_right:
        st.markdown("**Biggest Waste**")
        waste_brief = pd.concat([waste_kw_sheet, waste_st_sheet], ignore_index=True).drop_duplicates()
        if not waste_brief.empty:
            waste_brief = waste_brief.sort_values("spend", ascending=False).head(5)
            st.dataframe(waste_brief, use_container_width=True, hide_index=True)
        else:
            st.info("No major waste terms found.")

    with st.expander("Campaign Summary", expanded=False):
        if not campaign_view.empty:
            st.dataframe(campaign_view, use_container_width=True, hide_index=True)
        else:
            st.info("No campaign summary available.")
else:
    st.info("Upload your reports and click Audit My Account to begin.")
