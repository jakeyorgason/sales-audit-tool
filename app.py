import os
from typing import Any, Optional

import pandas as pd
import requests
import streamlit as st

from sales_audit_ingestion import SalesAuditEngine
from shared_ingestion_utils import to_excel_bytes_multi


st.set_page_config(
    page_title="Free Amazon Ads Audit",
    page_icon="📊",
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
    """
    <style>
        .main > div {
            padding-top: 1.1rem;
        }
        .block-container {
            padding-top: 1.1rem;
            padding-bottom: 1.75rem;
            max-width: 1440px;
        }
        [data-testid="stSidebar"] {
            background: #F3F4F6;
        }
        .brand-shell {
            background: linear-gradient(135deg, #EA580C 0%, #1F2937 100%);
            border-radius: 20px;
            padding: 50px 30px;
            color: white;
            margin-bottom: 1.25rem;
            box-shadow: 0 12px 28px rgba(0, 0, 0, 0.12);
        }
        .brand-title {
            font-size: 2.15rem;
            font-weight: 800;
            line-height: 1.05;
            margin: 0;
        }
        .brand-subtitle {
            font-size: 1rem;
            opacity: 0.96;
            margin-top: 0.55rem;
        }
        .section-title {
            font-size: 1.18rem;
            font-weight: 750;
            color: #111827;
            margin-bottom: 0.2rem;
        }
        .section-note {
            font-size: 0.92rem;
            color: #6b7280;
            margin-bottom: 0.8rem;
        }
        .metric-card {
            background: #ffffff;
            border: 1px solid #ececec;
            border-radius: 16px;
            padding: 14px 16px;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.04);
            min-height: 98px;
        }
        .metric-label {
            font-size: 0.75rem;
            color: #6B7280;
            margin-bottom: 0.25rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }
        .metric-value {
            font-size: 1.5rem;
            line-height: 1.1;
            font-weight: 800;
            color: #1F2937;
            word-break: break-word;
        }
        .metric-value.small {
            font-size: 1.1rem;
        }
        .metric-card.good { border-left: 4px solid #10B981; }
        .metric-card.warn { border-left: 4px solid #F59E0B; }
        .metric-card.bad { border-left: 4px solid #EF4444; }
        .metric-card.brand { border-left: 4px solid #F47322; }
        .status-pill {
            display: inline-block;
            padding: 8px 12px;
            border-radius: 999px;
            font-size: 0.82rem;
            font-weight: 700;
            margin-bottom: 0.65rem;
        }
        .status-pill.good {
            background: #ecfdf5;
            color: #047857;
            border: 1px solid #a7f3d0;
        }
        .status-pill.warn {
            background: #fffbeb;
            color: #b45309;
            border: 1px solid #fde68a;
        }
        .status-pill.bad {
            background: #fef2f2;
            color: #b91c1c;
            border: 1px solid #fecaca;
        }
        .summary-box {
            background: #ffffff;
            border: 1px solid #ececec;
            border-radius: 16px;
            padding: 16px 18px;
            margin-bottom: 10px;
        }
        .upload-note {
            color: #6b7280;
            font-size: 0.9rem;
            margin-top: -0.1rem;
            margin-bottom: 0.55rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown("## What to upload")
    st.markdown(
        """
- Bulk Sheet
- Sponsored Products Search Term Report
- Sponsored Products Targeting Report
- Sponsored Products Impression Share Report
- Sales & Traffic Business Report
- Sponsored Brands Campaign Report (optional)
"""
    )
    st.markdown("---")
    st.markdown("## Notes")
    st.markdown(
        """
- Keep the date ranges aligned where possible
- The SB report helps populate new-to-brand metrics
- Your uploaded data is only used to generate this audit
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
header_left, header_right = st.columns([1.35, 5.65], gap="medium")

with header_left:
    if logo_path:
        st.image(logo_path, width=500)

with header_right:
    st.markdown(
        """
        <div class="brand-shell">
            <div class="brand-title">Free Amazon Ads Audit</div>
            <div class="brand-subtitle">
                Upload your recent account reports and get a branded audit that highlights wasted spend, top opportunities, and winning terms.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# INPUTS
# =========================================================
st.markdown('<div class="section-title">Brand & Report Uploads</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-note">Enter your brand name, then upload the required files.</div>',
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
    f'<div class="upload-note">Status: {"Ready to run" if required_ready else "Please add your brand name and all 5 required reports."}</div>',
    unsafe_allow_html=True,
)

run_clicked = st.button(
    "Audit My Account",
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
    kw_winners = simplify_term_table(safe_df(winner_tables.get("keyword_winners")), "target").head(20)
    st_winners = simplify_term_table(safe_df(winner_tables.get("search_winners")), "customer_search_term").head(20)
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

    # TOP CONTENT
    top_content_left, top_content_right = st.columns([3.2, 1.35], gap="large")

    with top_content_left:
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

    with top_content_right:
        st.markdown("---")
        if not st.session_state["unlock_complete"]:
            st.markdown("### Unlock Your Branded Audit")
            st.markdown("Submit your details and we will generate your branded Google Sheets audit.")

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
                "Unlock My Audit",
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

                    with st.spinner("Generating your branded audit..."):
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
            st.success("Your branded audit is ready.")
            st.markdown(f"[Open Google Sheet]({created_report['url']})")

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

    export_sheets = {
        "KPI Summary": pd.DataFrame([kpis]),
        "Waste Summary": pd.DataFrame([waste_summary]),
        "Campaign Summary": campaign_view,
        "Keyword Spend": top_kw_sheet,
        "Search Term Spend": top_st_sheet,
        "KW Waste": waste_kw_sheet,
        "ST Waste": waste_st_sheet,
        "KW Winners": kw_winners_sheet,
        "ST Winners": st_winners_sheet,
    }

    export_bytes = to_excel_bytes_multi(export_sheets)
    st.download_button(
        label="Download Sales Audit Workbook",
        data=export_bytes,
        file_name=f"{brand_name or 'sales_audit'}_audit_workbook.xlsx".replace(" ", "_"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("Upload your reports and click Audit My Account to begin.")
