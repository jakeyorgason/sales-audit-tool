import os
import base64
from typing import Any, Optional

import pandas as pd
import requests
import streamlit as st

from sales_audit_ingestion import SalesAuditEngine


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Free Amazon Ads Audit | Evolved Commerce",
    page_icon="assets/ec_logo2.jpg",
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


def load_banner_path() -> Optional[str]:
    possible_paths = [
        "assets/ec_banner.png",
        "assets/ec_banner.jpg",
        "assets/banner.png",
        "ec_banner.png",
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None


def image_to_base64_src(path: Optional[str]) -> str:
    if not path or not os.path.exists(path):
        return ""

    ext = os.path.splitext(path)[1].lower().replace(".", "")
    mime = "image/png" if ext == "png" else "image/jpeg"

    with open(path, "rb") as image_file:
        encoded = base64.b64encode(image_file.read()).decode("utf-8")

    return f"data:{mime};base64,{encoded}"


def format_currency(value: float) -> str:
    return f"${get_number(value):,.2f}"


def format_percent(value: float) -> str:
    return f"{get_number(value):,.2f}%"


def format_number(value: float) -> str:
    return f"{get_number(value):,.2f}"


def render_metric_card(label: str, value: str, tone: str = "brand", small: bool = False) -> None:
    value_class = "metric-value small" if small else "metric-value"
    st.markdown(
        f"""
        <div class="metric-card {tone}">
            <div class="metric-label">{label}</div>
            <div class="{value_class}">{value}</div>
        </div>
        """,
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

    out = out.rename(
        columns={
            term_col: "term",
            "spend": "spend",
            "sales": "sales",
            "acos_pct": "acos",
        }
    )

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
    """
    Build raw numeric rows for the Google Sheet payload.

    NOTE:
    For this audit workflow, the term-level spend/sales tables are arriving in cents,
    while the KPI summary is already in dollars. We normalize to dollars here.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    if "acos" in out.columns and "acos_pct" not in out.columns:
        out["acos_pct"] = out["acos"] * 100

    keep_cols = [c for c in [term_col, "spend", "sales", "acos_pct"] if c in out.columns]
    out = out[keep_cols].copy()

    out = out.rename(
        columns={
            term_col: "term",
            "spend": "spend",
            "sales": "sales",
            "acos_pct": "acos",
        }
    )

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
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

        :root {
            --ec-orange: #ff6a00;
            --ec-orange-dark: #f26300;
            --ec-orange-soft: #fff1e7;
            --ec-black: #1f2833;
            --ec-charcoal: #26313d;
            --ec-muted: #5d6670;
            --ec-border: #ded8d0;
            --ec-bg: #f4f1ec;
            --ec-card: #ffffff;
            --ec-green: #16a34a;
            --ec-yellow: #d97706;
            --ec-red: #dc2626;
            --ec-hero-gray: #d3d7d4;
            --ec-navy: #202b36;
        }

        html, body, [class*="css"] {
            font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }

        .stApp {
            background: var(--ec-navy);
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

        [data-testid="stSidebar"] {
            background: #1f2833;
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

        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] li {
            font-size: 0.91rem;
            line-height: 1.5;
        }

        [data-testid="stSidebar"] .stButton > button {
            background: var(--ec-orange) !important;
            border-color: var(--ec-orange) !important;
            color: #ffffff !important;
            box-shadow: 0 14px 28px rgba(255, 106, 0, 0.22);
        }

        .site-hero {
            background: var(--ec-hero-gray);
            border-radius: 34px;
            padding: 30px 34px 0 34px;
            margin: 0 auto 1.35rem auto;
            min-height: 525px;
            box-shadow: 0 30px 90px rgba(17, 24, 39, 0.22);
            position: relative;
            overflow: hidden;
            border: 1px solid rgba(17,17,17,0.08);
        }

        .site-hero-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 24px;
            position: relative;
            z-index: 4;
        }

        .hero-logo {
            width: 245px;
            max-width: 42vw;
            height: auto;
            display: block;
        }

        .hero-banner-wrap {
            display: flex;
            justify-content: flex-end;
            align-items: center;
        }

        .hero-banner-logo {
            width: 360px;
            max-width: 36vw;
            height: auto;
            display: block;
        }

        .site-hero-content {
            display: grid;
            grid-template-columns: 1.05fr 0.95fr;
            gap: 24px;
            align-items: center;
            min-height: 440px;
            position: relative;
            z-index: 2;
        }

        .site-hero-copy {
            padding: 38px 0 48px 8px;
        }

        .hero-kicker {
            display: inline-flex;
            align-items: center;
            padding: 8px 14px;
            background: rgba(255, 106, 0, 0.12);
            border: 1px solid rgba(255, 106, 0, 0.24);
            border-radius: 999px;
            color: var(--ec-orange-dark);
            font-size: 0.76rem;
            font-weight: 900;
            letter-spacing: 0.075em;
            text-transform: uppercase;
            margin-bottom: 18px;
        }

        .brand-title {
            font-size: clamp(3rem, 5.8vw, 6.1rem);
            font-weight: 900;
            letter-spacing: -0.078em;
            line-height: 0.88;
            margin: 0;
            color: var(--ec-black);
            max-width: 760px;
        }

        .brand-title .accent {
            color: var(--ec-orange);
            display: inline-block;
        }

        .brand-subtitle {
            font-size: 1.05rem;
            color: #3f4852;
            margin-top: 1.2rem;
            max-width: 690px;
            line-height: 1.62;
            font-weight: 600;
        }

        .hero-proof-row {
            display: flex;
            flex-wrap: wrap;
            gap: 9px;
            margin-top: 26px;
        }

        .hero-proof-item {
            background: rgba(31,40,51,0.92);
            color: #ffffff;
            border-radius: 999px;
            padding: 10px 13px;
            font-size: 0.82rem;
            font-weight: 850;
        }

        .site-hero-visual {
            position: relative;
            min-height: 395px;
        }

        .orange-shape {
            position: absolute;
            width: 520px;
            height: 520px;
            border-radius: 46% 54% 50% 50%;
            background: linear-gradient(135deg, #ff7a00 0%, #f26300 48%, #ffb347 100%);
            right: -90px;
            bottom: -160px;
            transform: rotate(-18deg);
            box-shadow: 0 30px 90px rgba(255, 106, 0, 0.28);
        }

        .audit-card {
            position: absolute;
            background: rgba(255,255,255,0.94);
            border: 1px solid rgba(17,17,17,0.08);
            border-radius: 18px;
            padding: 15px 16px;
            width: 220px;
            box-shadow: 0 20px 50px rgba(17, 24, 39, 0.16);
            z-index: 3;
        }

        .audit-card strong {
            display: block;
            color: var(--ec-black);
            font-weight: 900;
            letter-spacing: -0.02em;
            margin-bottom: 5px;
        }

        .audit-card span {
            display: block;
            color: #5b6470;
            font-size: 0.86rem;
            line-height: 1.38;
            font-weight: 600;
        }

        .floating-one {
            right: 235px;
            top: 94px;
        }

        .floating-two {
            right: 70px;
            top: 215px;
        }

        .partner-badge {
            position: absolute;
            right: 20px;
            bottom: 34px;
            background: #ffffff;
            border-radius: 16px;
            padding: 13px 16px;
            color: var(--ec-black);
            font-size: 0.86rem;
            line-height: 1.15;
            font-weight: 800;
            z-index: 4;
            box-shadow: 0 16px 42px rgba(17, 24, 39, 0.16);
        }

        .partner-badge strong {
            color: var(--ec-orange);
        }

        .trust-strip {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 14px;
            margin: 0.4rem 0 1.2rem 0;
        }

        .trust-card {
            background: rgba(255,255,255,0.92);
            border: 1px solid var(--ec-border);
            border-radius: 22px;
            padding: 17px 18px;
            box-shadow: 0 14px 36px rgba(17,17,17,0.08);
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

        .audit-form-card {
            background: #f4f1ec;
            border: 1px solid rgba(255,255,255,0.14);
            border-radius: 30px;
            padding: 30px 34px 34px 34px;
            margin-top: 1.25rem;
            box-shadow: 0 26px 80px rgba(17, 24, 39, 0.18);
        }

        .audit-form-title {
            font-size: 2.05rem;
            line-height: 1;
            letter-spacing: -0.055em;
            font-weight: 900;
            color: #1f2833;
            margin: 0 0 0.45rem 0;
        }

        .audit-form-title:before {
            content: "";
            display: inline-block;
            width: 11px;
            height: 11px;
            background: #ff6a00;
            border-radius: 999px;
            margin-right: 11px;
            transform: translateY(-3px);
        }

        .audit-form-subtitle {
            color: #5d6670;
            font-size: 1rem;
            line-height: 1.55;
            font-weight: 600;
            max-width: 860px;
        }

        .required-pill-row {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 14px;
        }

        .required-pill {
            background: #ffffff;
            border: 1px solid #ded8d0;
            color: #1f2833;
            border-radius: 999px;
            padding: 7px 11px;
            font-size: 0.78rem;
            font-weight: 800;
        }

        .brand-entry-card {
            display: grid;
            grid-template-columns: 1.35fr 1fr;
            gap: 24px;
            background: #f4f1ec;
            border: 1px solid #ded8d0;
            border-radius: 28px;
            padding: 26px;
            margin: 1.25rem 0 0.75rem 0;
            box-shadow: 0 18px 48px rgba(17,24,39,0.16);
        }

        .brand-entry-left,
        .brand-entry-right {
            background: #ffffff;
            border: 1px solid rgba(222,216,208,0.95);
            border-radius: 22px;
            padding: 22px 24px;
            box-shadow: 0 10px 26px rgba(17,24,39,0.04);
        }

        .brand-entry-left,
        .brand-entry-right {
            min-geight: 150px;
        }

        .brand-step-eyebrow {
            display: inline-flex;
            align-items: center;
            background: #fff1e7;
            color: #ff6a00;
            border: 1px solid rgba(255,106,0,0.22);
            border-radius: 999px;
            padding: 6px 10px;
            font-size: 0.74rem;
            font-weight: 900;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 10px;
        }

        .brand-step-title {
            color: #1f2833;
            font-size: 1.45rem;
            line-height: 1.05;
            font-weight: 900;
            letter-spacing: -0.04em;
            margin-bottom: 6px;
        }

        .brand-step-title.small {
            font-size: 1.05rem;
        }

        .brand-step-copy {
            color: #5d6670;
            font-size: 0.95rem;
            line-height: 1.45;
            font-weight: 650;
        }

        .dark-section-heading {
            color: #ffffff;
            font-size: 1.2rem;
            font-weight: 900;
            letter-spacing: -0.025em;
            margin: 1.7rem 0 0.85rem 0;
        }

        .dark-section-heading:before {
            content: "";
            display: inline-block;
            width: 9px;
            height: 9px;
            background: #ff6a00;
            border-radius: 999px;
            margin-right: 9px;
            transform: translateY(-1px);
        }

        .section-title {
            font-size: 1.32rem;
            font-weight: 900;
            letter-spacing: -0.035em;
            color: var(--ec-black);
            margin-bottom: 0.25rem;
            margin-top: 1.25rem;
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

        .metric-card {
            background: var(--ec-card);
            border: 1px solid var(--ec-border);
            border-radius: 22px;
            padding: 17px 18px;
            box-shadow: 0 14px 38px rgba(17,17,17,0.07);
            min-height: 104px;
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
            box-shadow: 0 14px 38px rgba(17,17,17,0.07);
            color: var(--ec-charcoal);
            line-height: 1.62;
        }

        .summary-box strong {
            color: var(--ec-black);
            font-weight: 900;
        }

        .upload-note {
            background: #ffffff;
            border: 1px solid #ded8d0;
            border-radius: 16px;
            padding: 13px 15px;
            color: #5d6670;
            font-size: 0.94rem;
            margin-top: 0.4rem;
            margin-bottom: 0.8rem;
            font-weight: 800;
        }

        .lead-card {
            background:
                radial-gradient(circle at top right, rgba(255,106,0,0.18), transparent 20rem),
                var(--ec-black);
            color: #ffffff;
            border-radius: 30px;
            padding: 34px 36px;
            box-shadow: 0 28px 80px rgba(17,17,17,0.20);
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

        div[data-testid="stTextInput"] {
            max-width: 100%;
        }

        div[data-testid="stTextInput"] label {
            font-weight: 900 !important;
            color: #1f2833 !important;
            margin-bottom: 0.45rem !important;
        }

        .stTextInput input {
            border-radius: 18px !important;
            border: 1px solid #d9d3ca !important;
            min-height: 58px !important;
            height: 58px !important;
            padding: 0 18px !important;
            font-size: 1rem !important;
            background: #ffffff !important;
            color: #1f2833 !important;
            box-shadow: 0 10px 26px rgba(17,24,39,0.08) !important;
        }

        .stTextInput input::placeholder {
            color: #6b7280 !important;
            opacity: 1 !important;
        }

        .stTextInput input:focus {
            border-color: #ff6a00 !important;
            box-shadow: 0 0 0 4px rgba(255,106,0,0.16), 0 10px 26px rgba(17,24,39,0.08) !important;
        }

        /* The Amazon brand input */
        div[data-testid="stTextInput"]:has(input[aria-label="Amazon Brand Name"]) {
            margin-top: 0 !important;
            margin-bottom: 1.5rem !important;
        }
        
        div[data-testid="stTextInput"]:has(input[aria-label="Amazon Brand Name"]) > div {
            min-height: 50px !important;
        }

        div[data-testid="stTextInput"]:has(input[aria-label="Amazon Brand Name"]) input {
            background: #ffffff !important;
            color: #1f2833 !important;
            border: 2px solid rgba(255,106,0,0.60) !important;
            border-radius: 16px !important;
        
            min-height: 48px !important;
            height: 48px !important;
        
            padding: 0 20px !important;
            font-size: 1rem !important;
            font-weight: 750 !important;
            line-height: 48px !important;
        
            box-shadow: 0 10px 24px rgba(0,0,0,0.14) !important;
            overflow: visible !important;
        }

        div[data-testid="stTextInput"]:has(input[aria-label="Amazon Brand Name"]) input::placeholder {
            color: #374151 !important;
            opacity: 1 !important;
            font-weight: 700 !important;
        }

        div[data-testid="stTextInput"]:has(input[aria-label="Amazon Brand Name"]) input:focus {
            border-color: #ff6a00 !important;
            box-shadow: 0 0 0 4px rgba(255,106,0,0.22), 0 14px 34px rgba(0,0,0,0.16) !important;
        }

        [data-testid="stFileUploader"] {
            background: #ffffff;
            border: 1px solid #ded8d0;
            border-radius: 22px;
            padding: 0.85rem;
            box-shadow: 0 14px 34px rgba(17,17,17,0.065);
        }

        [data-testid="stFileUploader"]:hover {
            border-color: rgba(255,106,0,0.55);
            box-shadow: 0 18px 42px rgba(17,17,17,0.09);
        }

        [data-testid="stFileUploader"] section {
            background: #fbfaf7 !important;
            border: 1px dashed rgba(255,106,0,0.45) !important;
            border-radius: 18px !important;
            padding: 18px !important;
        }

        [data-testid="stFileUploader"] label {
            color: #1f2833 !important;
            font-weight: 900 !important;
            font-size: 0.9rem !important;
        }

        [data-testid="stFileUploader"] small {
            color: #5d6670 !important;
            font-weight: 600 !important;
        }

        [data-testid="stFileUploader"] button {
            background: #ff6a00 !important;
            color: #ffffff !important;
            border: 1px solid #ff6a00 !important;
            border-radius: 999px !important;
            font-weight: 900 !important;
            box-shadow: 0 10px 24px rgba(255,106,0,0.22) !important;
        }

        [data-testid="stFileUploader"] button:hover {
            background: #f26300 !important;
            border-color: #f26300 !important;
            color: #ffffff !important;
        }

        [data-testid="stFileUploader"] button * {
            color: #ffffff !important;
        }

        .stButton > button {
            border-radius: 999px !important;
            border: 1px solid var(--ec-black) !important;
            background: var(--ec-black) !important;
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
            box-shadow: 0 18px 38px rgba(255,106,0,0.24);
        }

        .stButton > button[kind="primary"],
        button[kind="primary"] {
            background: #ff6a00 !important;
            border-color: #ff6a00 !important;
            color: #ffffff !important;
            border-radius: 999px !important;
            font-weight: 900 !important;
            min-height: 54px !important;
            box-shadow: 0 16px 34px rgba(255,106,0,0.25) !important;
        }

        .stButton > button[kind="primary"]:hover,
        button[kind="primary"]:hover {
            background: #f26300 !important;
            border-color: #f26300 !important;
        }

        .stButton > button:disabled {
            background: #d7cec4 !important;
            border-color: #d7cec4 !important;
            color: #ffffff !important;
            box-shadow: none !important;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--ec-border);
            border-radius: 18px;
            overflow: hidden;
            box-shadow: 0 12px 34px rgba(17,17,17,0.06);
        }

        div[data-testid="stExpander"] {
            background: rgba(255,255,255,0.92);
            border: 1px solid var(--ec-border);
            border-radius: 18px;
            box-shadow: 0 10px 28px rgba(17,17,17,0.05);
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

        @media (max-width: 1000px) {
            .site-hero-content {
                grid-template-columns: 1fr;
            }

            .site-hero-visual {
                display: none;
            }

            .site-hero {
                min-height: auto;
                padding-bottom: 34px;
            }

            .hero-banner-wrap {
                display: none;
            }

            .trust-strip {
                grid-template-columns: 1fr;
            }

            .brand-title {
                font-size: 3.1rem;
            }

            .brand-entry-card {
                grid-template-columns: 1fr;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown("## Free Audit Checklist")
    st.markdown(
        """
Upload these reports to generate your audit:

- Bulk Sheet
- Sponsored Products Search Term Report
- Sponsored Products Targeting Report
- Sponsored Products Impression Share Report
- Sales & Traffic Business Report
- Sponsored Brands Campaign Report *(optional)*
"""
    )

    st.markdown("---")

    st.markdown("## What You’ll See")
    st.markdown(
        """
- Account health verdict
- Executive KPI snapshot
- Wasted spend indicators
- Top keywords and search terms
- Campaign-level performance summary
- Personalized Google Sheets audit
"""
    )

    st.markdown("---")

    st.markdown("## Best Results")
    st.markdown(
        """
Use matching date ranges where possible. The Sponsored Brands report helps populate new-to-brand context, but it is not required.
"""
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
logo_src = image_to_base64_src(logo_path)

banner_path = load_banner_path()
banner_src = image_to_base64_src(banner_path)

logo_html = ""
if logo_src:
    logo_html = f'<img src="{logo_src}" class="hero-logo" alt="Evolved Commerce logo">'

banner_html = ""
if banner_src:
    banner_html = f'<img src="{banner_src}" class="hero-banner-logo" alt="Evolved Commerce">'

hero_html = (
    '<div class="site-hero">'
        '<div class="site-hero-top">'
            f'{logo_html}'
            '<div class="hero-banner-wrap">'
                f'{banner_html}'
            '</div>'
        '</div>'
        '<div class="site-hero-content">'
            '<div class="site-hero-copy">'
                '<div class="hero-kicker">Free Amazon Advertising Audit</div>'
                '<div class="brand-title">'
                    'Your Free<br>'
                    '<span class="accent">Amazon Ads Audit</span>'
                '</div>'
                '<div class="brand-subtitle">'
                    'Upload your recent Amazon reports and uncover wasted spend, campaign inefficiencies, '
                    'winning search terms, and growth opportunities your brand can act on.'
                '</div>'
                '<div class="hero-proof-row">'
                    '<span class="hero-proof-item">Wasted spend review</span>'
                    '<span class="hero-proof-item">Campaign health snapshot</span>'
                    '<span class="hero-proof-item">Search term opportunities</span>'
                '</div>'
            '</div>'
            '<div class="site-hero-visual">'
                '<div class="orange-shape"></div>'
                '<div class="audit-card floating-one">'
                    '<strong>Wasted Spend</strong>'
                    '<span>Find inefficient clicks and high ACOS terms.</span>'
                '</div>'
                '<div class="audit-card floating-two">'
                    '<strong>Growth Terms</strong>'
                    '<span>Surface the search terms driving sales.</span>'
                '</div>'
                '<div class="partner-badge">'
                    'amazon ads<br><strong>audit ready</strong>'
                '</div>'
            '</div>'
        '</div>'
    '</div>'
    '<div class="trust-strip">'
        '<div class="trust-card">'
            '<strong>Built for Amazon brands</strong>'
            '<span>Review ads performance, search terms, targeting, and account efficiency in one guided workflow.</span>'
        '</div>'
        '<div class="trust-card">'
            '<strong>Focused on profitable growth</strong>'
            '<span>Spot wasted spend, identify winners, and understand where advertising is helping or hurting sales velocity.</span>'
        '</div>'
        '<div class="trust-card">'
            '<strong>Personalized report output</strong>'
            '<span>Submit your details after the preview to generate a shareable Google Sheets audit for your brand.</span>'
        '</div>'
    '</div>'
)

st.markdown(hero_html, unsafe_allow_html=True)


# =========================================================
# INPUTS
# =========================================================
st.markdown(
    """
    <div class="audit-form-card">
        <div class="audit-form-title">Start Your Free Audit</div>
        <div class="audit-form-subtitle">
            Add your brand name and upload the five required Amazon reports. We’ll use them to find wasted spend,
            winning terms, campaign issues, and growth opportunities.
        </div>
        <div class="required-pill-row">
            <span class="required-pill">Bulk Sheet</span>
            <span class="required-pill">SP Search Terms</span>
            <span class="required-pill">SP Targeting</span>
            <span class="required-pill">SP Impression Share</span>
            <span class="required-pill">Sales & Traffic</span>
            <span class="required-pill">SB Optional</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="brand-entry-card">
        <div class="brand-entry-left">
            <div class="brand-step-eyebrow">Step 1</div>
            <div class="brand-step-title">Enter your Amazon brand name</div>
            <div class="brand-step-copy">
                This name will be used in your audit preview and final report.
            </div>
        </div>
        <div class="brand-entry-right">
            <div class="brand-step-eyebrow">Step 2</div>
            <div class="brand-step-title small">Then upload your reports</div>
            <div class="brand-step-copy">
                Once the five required files are added, the audit button will unlock and generate your preview.
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

brand_name_input = st.text_input(
    "Amazon Brand Name",
    value=st.session_state.get("lead_brand_name", ""),
    placeholder="Example: Evolved Commerce",
    key="sales_audit_brand_name",
    label_visibility="collapsed",
)

st.markdown(
    '<div class="dark-section-heading">Required Reports</div>',
    unsafe_allow_html=True,
)

u1, u2 = st.columns(2, gap="large")
u3, u4 = st.columns(2, gap="large")
u5, u6 = st.columns(2, gap="large")

with u1:
    bulk_file = st.file_uploader(
        "1. Bulk Sheet",
        type=["xlsx", "xls", "csv"],
        key="sales_audit_bulk_file",
    )

with u2:
    impression_share_file = st.file_uploader(
        "2. SP Impression Share Report",
        type=["csv", "xlsx", "xls"],
        key="sales_audit_impression_share_file",
    )

with u3:
    targeting_file = st.file_uploader(
        "3. SP Targeting Report",
        type=["csv", "xlsx", "xls"],
        key="sales_audit_targeting_file",
    )

with u4:
    search_term_file = st.file_uploader(
        "4. SP Search Term Report",
        type=["csv", "xlsx", "xls"],
        key="sales_audit_search_term_file",
    )

with u5:
    business_report_file = st.file_uploader(
        "5. Sales & Traffic Business Report",
        type=["csv", "xlsx", "xls"],
        key="sales_audit_business_report_file",
    )

with u6:
    sb_campaign_file = st.file_uploader(
        "Optional: Sponsored Brands Campaign Report",
        type=["csv", "xlsx", "xls"],
        key="sales_audit_sb_campaign_file",
    )

required_ready = all(
    [
        brand_name_input.strip(),
        bulk_file is not None,
        impression_share_file is not None,
        targeting_file is not None,
        search_term_file is not None,
        business_report_file is not None,
    ]
)

status_text = (
    "Ready to generate your audit."
    if required_ready
    else "Add your brand name and all five required reports to unlock the audit button."
)

st.markdown(f'<div class="upload-note">{status_text}</div>', unsafe_allow_html=True)

run_clicked = st.button(
    "Generate My Free Audit",
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


# =========================================================
# RESULTS
# =========================================================
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

    top_kw = simplify_term_table(keyword_spend_table, "target").head(20)
    top_st = simplify_term_table(search_term_spend_table, "customer_search_term").head(20)
    campaign_view = simplify_campaign_table(campaign_summary).head(20)

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
        f"""
        <div class="summary-box">
            <strong>Summary:</strong> {health_summary.get("summary", "No summary available.")}
            <br><br>
            <strong>Narrative:</strong> {narrative}
        </div>
        """,
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

    st.markdown(
        """
        <div class="lead-card">
            <span class="mini">Your personalized report is ready to build</span>
            <h2>Get your free Amazon Ads audit report.</h2>
            <p>
                Submit your details and we’ll generate a personalized Google Sheets audit
                with your KPI snapshot, wasted spend review, winning terms, and campaign summary.
            </p>
        </div>
        """,
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
        report_url = created_report.get("url", "")
        st.success("Your personalized audit is ready.")
        st.markdown(
            f"""
            <div class="summary-box">
                <strong>Your Google Sheets audit has been created.</strong><br>
                <a href="{report_url}" target="_blank">Open your personalized Amazon Ads audit →</a>
            </div>
            """,
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
    st.info("Upload your reports and click Generate My Free Audit to begin.")


# =========================================================
# FOOTER
# =========================================================
footer_html = (
    '<div style="'
        'text-align:center;'
        'color:#cfd6dd;'
        'font-size:0.86rem;'
        'font-weight:600;'
        'margin-top:2rem;'
        'padding-top:1rem;'
        'border-top:1px solid rgba(255,255,255,0.12);'
    '">'
        'Free Amazon Ads Audit by Evolved Commerce'
    '</div>'
)

st.markdown(footer_html, unsafe_allow_html=True)
