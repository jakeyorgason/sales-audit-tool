import os
from typing import Any, Optional

import pandas as pd
import requests
import streamlit as st

from sales_audit_ingestion import SalesAuditEngine


st.set_page_config(
    page_title="Amazon Ads Audit | Evolved Commerce",
    page_icon="📊",
    layout="wide",
)


# =========================================================
# SETTINGS / HELPERS
# =========================================================
def get_setting(name: str, default: Optional[str] = None) -> Optional[str]:
    if name in os.environ:
        return os.environ.get(name, default)
    try:
        return st.secrets[name]
    except Exception:
        return default


def safe_dict(value: Any) -> dict:
    return value if isinstance(value, dict) else {}


def safe_df(value: Any) -> pd.DataFrame:
    return value if isinstance(value, pd.DataFrame) else pd.DataFrame()


def get_number(value: Any, default: float = 0.0) -> float:
    try:
        if value in (None, ""):
            return default
        return float(value)
    except Exception:
        return default


def format_currency(value: Any) -> str:
    return f"${get_number(value):,.2f}"


def format_percent(value: Any) -> str:
    return f"{get_number(value):,.2f}%"


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


def simplify_term_table(df: pd.DataFrame, term_col: str, limit: int = 10) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()
    if "acos" in out.columns and "acos_pct" not in out.columns:
        out["acos_pct"] = pd.to_numeric(out["acos"], errors="coerce").fillna(0) * 100

    keep_cols = [c for c in [term_col, "spend", "sales", "acos_pct"] if c in out.columns]
    out = out[keep_cols].copy()
    out = out.rename(columns={term_col: "term", "acos_pct": "acos"})

    for col in ["spend", "sales", "acos"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    out = out.sort_values(["spend", "sales"], ascending=[False, False]).head(limit).reset_index(drop=True)
    out["spend"] = out["spend"].map(lambda x: f"${x:,.2f}")
    out["sales"] = out["sales"].map(lambda x: f"${x:,.2f}")
    out["acos"] = out["acos"].map(lambda x: f"{x:,.2f}%")
    return out


def simplify_raw_for_sheet(df: pd.DataFrame, term_col: str, limit: int = 20) -> list[dict]:
    if df is None or df.empty:
        return []

    out = df.copy()
    if "acos" in out.columns and "acos_pct" not in out.columns:
        out["acos_pct"] = pd.to_numeric(out["acos"], errors="coerce").fillna(0) * 100

    keep_cols = [c for c in [term_col, "spend", "sales", "acos_pct"] if c in out.columns]
    out = out[keep_cols].copy().rename(columns={term_col: "term", "acos_pct": "acos"})

    for col in ["spend", "sales", "acos"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    out = out.sort_values(["spend", "sales"], ascending=[False, False]).head(limit).reset_index(drop=True)
    return out.to_dict("records")


def normalize_records_for_sheet(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []

    out = df.copy()
    out = out.where(pd.notnull(out), None)

    for col in out.columns:
        col_l = str(col).lower().strip()
        if col_l in {"spend", "sales", "ad_sales", "total_sales", "organic_sales", "ntb_sales", "wasted_spend", "spend_no_sale"}:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).round(2)
        elif col_l in {"impressions", "clicks", "orders", "units_ordered", "sessions", "ntb_orders"}:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).round(0).astype(int)
        elif col_l in {"ctr", "cvr", "acos", "acos_pct", "impression_share_pct", "tacos_pct", "organic_share_pct", "unit_session_percentage", "ntb_sales_pct", "ntb_orders_pct", "wasted_spend_pct"}:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).round(2)
        elif col_l == "roas":
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).round(2)
        elif col_l == "cpc":
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).round(2)

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
    webhook_url = get_setting("GOOGLE_SHEET_WEBHOOK_URL")
    template_id = get_setting("GOOGLE_SHEETS_TEMPLATE_ID")
    destination_folder_id = get_setting("GOOGLE_DRIVE_FOLDER_ID")

    if not webhook_url or not template_id or not destination_folder_id:
        raise RuntimeError("Google Sheet template or destination folder ID is not configured.")

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


def send_lead_webhook(lead_payload: dict) -> None:
    webhook_url = get_setting("LEAD_WEBHOOK_URL")
    if not webhook_url:
        return
    response = requests.post(webhook_url, json=lead_payload, timeout=30)
    response.raise_for_status()


# =========================================================
# STYLING
# =========================================================
st.markdown(
    """
    <style>
        .main > div { padding-top: 1.0rem; }
        .block-container { padding-top: 1rem; padding-bottom: 1.75rem; max-width: 1440px; }
        [data-testid="stSidebar"] { background: #F3F4F6; }
        .hero { background: linear-gradient(135deg, #EA580C 0%, #1F2937 100%); border-radius: 20px; padding: 28px 28px; color: white; margin-bottom: 1rem; }
        .hero h1 { margin: 0; font-size: 2rem; }
        .hero p { margin: .4rem 0 0 0; opacity: .96; }
        .card-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin: 12px 0 18px; }
        .mini-card { border:1px solid #ead8c2; border-radius:14px; padding:14px 16px; background:#fcf7f1; }
        .section-title { font-size:1.2rem; font-weight:800; margin:.8rem 0 .25rem; }
        .section-note { color:#6b7280; margin-bottom:.75rem; }
        .metric { background:#fff; border:1px solid #ececec; border-radius:16px; padding:14px 16px; }
        .metric-label { font-size:.75rem; font-weight:700; color:#6b7280; text-transform:uppercase; letter-spacing:.04em; }
        .metric-value { font-size:1.4rem; font-weight:800; color:#1f2937; }
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
        for key in [
            "sales_audit_results",
            "sales_audit_unlocked",
            "sales_audit_created_report",
            "sales_audit_error",
        ]:
            st.session_state.pop(key, None)
        st.rerun()


# =========================================================
# HEADER
# =========================================================
logo_path = load_logo_path()
left, right = st.columns([1.25, 4.75], gap="medium")
with left:
    if logo_path:
        st.image(logo_path, width=260)
with right:
    st.markdown(
        """
        <div class="hero">
            <h1>Free Amazon Ads Audit | Evolved Commerce</h1>
            <p>Upload your recent Amazon reports and get a branded audit that highlights wasted spend, top opportunities, and winning terms.</p>
        </div>
        <div class="card-grid">
            <div class="mini-card"><strong>1. Upload 5 reports</strong><br><span style="color:#6b7280">Recent reporting windows work best.</span></div>
            <div class="mini-card"><strong>2. Run your audit</strong><br><span style="color:#6b7280">We process the reports and prepare your branded analysis.</span></div>
            <div class="mini-card"><strong>3. Unlock results</strong><br><span style="color:#6b7280">Submit your details to receive the completed audit.</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# INPUTS
# =========================================================
st.markdown('<div class="section-title">Brand & Report Uploads</div>', unsafe_allow_html=True)
st.markdown('<div class="section-note">Enter your brand name, then upload the required files.</div>', unsafe_allow_html=True)

brand_name = st.text_input("Brand Name", placeholder="Your brand name", key="sales_audit_brand_name")

c1, c2 = st.columns(2)
c3, c4 = st.columns(2)
c5, c6 = st.columns(2)
with c1:
    bulk_file = st.file_uploader("Bulk Sheet", type=["xlsx", "xls", "csv"], key="bulk_file")
with c2:
    impression_share_file = st.file_uploader("SP Impression Share Report", type=["xlsx", "xls", "csv"], key="impression_share_file")
with c3:
    targeting_file = st.file_uploader("SP Targeting Report", type=["xlsx", "xls", "csv"], key="targeting_file")
with c4:
    search_term_file = st.file_uploader("SP Search Term Report", type=["xlsx", "xls", "csv"], key="search_term_file")
with c5:
    business_report_file = st.file_uploader("Sales & Traffic Business Report", type=["xlsx", "xls", "csv"], key="business_report_file")
with c6:
    sb_campaign_file = st.file_uploader("Sponsored Brands Campaign Report (optional)", type=["xlsx", "xls", "csv"], key="sb_campaign_file")

required_ready = all([
    brand_name.strip(),
    bulk_file is not None,
    impression_share_file is not None,
    targeting_file is not None,
    search_term_file is not None,
    business_report_file is not None,
])

st.caption("Ready to run" if required_ready else "Please add your brand name and all 5 required reports.")

if st.button("Audit My Account", type="primary", use_container_width=True, disabled=not required_ready):
    st.session_state.pop("sales_audit_error", None)
    st.session_state["sales_audit_unlocked"] = False
    st.session_state["sales_audit_created_report"] = None
    with st.spinner("Running your audit..."):
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
            engine.brand_name = brand_name.strip()
            results = engine.process()
            results["brand_name"] = brand_name.strip()
            results["has_sb_report"] = sb_campaign_file is not None
            st.session_state["sales_audit_results"] = results
        except Exception as exc:
            st.session_state["sales_audit_error"] = f"We hit an issue while processing your files: {exc}"

results = safe_dict(st.session_state.get("sales_audit_results", {}))
error_message = st.session_state.get("sales_audit_error")
if error_message:
    st.error(error_message)

if results and not st.session_state.get("sales_audit_unlocked", False):
    st.markdown('<div class="section-title">Unlock Your Branded Audit</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-note">Submit your details and we will generate your branded Google Sheets audit.</div>', unsafe_allow_html=True)

    with st.form("unlock_audit_form"):
        l1, l2 = st.columns(2)
        with l1:
            lead_name = st.text_input("Full Name", value="", key="lead_name")
            lead_email = st.text_input("Work Email", value="", key="lead_email")
        with l2:
            lead_phone = st.text_input("Phone Number", value="", key="lead_phone")
            lead_brand = st.text_input("Brand Name", value=results.get("brand_name", ""), key="lead_brand")
        submitted = st.form_submit_button("Unlock My Audit", type="primary", use_container_width=True)

    if submitted:
        if not all([lead_name.strip(), lead_email.strip(), lead_phone.strip(), lead_brand.strip()]):
            st.error("Please complete all four fields before unlocking your audit.")
        else:
            with st.spinner("Generating your branded audit..."):
                try:
                    kpis = safe_dict(results.get("kpi_summary"))
                    waste_summary = safe_dict(results.get("waste_summary"))
                    keyword_spend_table = safe_df(results.get("keyword_spend_table"))
                    search_term_spend_table = safe_df(results.get("search_term_spend_table"))
                    waste_tables = safe_dict(results.get("waste_tables"))
                    winner_tables = safe_dict(results.get("winner_tables"))
                    campaign_summary = safe_df(results.get("campaign_summary"))

                    kw_zero = safe_df(waste_tables.get("keyword_zero_sale"))
                    kw_high = safe_df(waste_tables.get("keyword_high_acos"))
                    st_zero = safe_df(waste_tables.get("search_zero_sale"))
                    st_high = safe_df(waste_tables.get("search_high_acos"))
                    kw_winners_raw = safe_df(winner_tables.get("keyword_winners"))
                    st_winners_raw = safe_df(winner_tables.get("search_winners"))

                    waste_kw_raw = pd.concat([kw_zero, kw_high], ignore_index=True).drop_duplicates() if (not kw_zero.empty or not kw_high.empty) else pd.DataFrame()
                    waste_st_raw = pd.concat([st_zero, st_high], ignore_index=True).drop_duplicates() if (not st_zero.empty or not st_high.empty) else pd.DataFrame()

                    report_name = f"{lead_brand.strip()} - Amazon Ads Audit"
                    date_range_label = (results.get("date_range_label") or "").strip() or "MM/DD - MM/DD"

                    created_report = create_google_sheet_report(
                        brand_name=lead_brand.strip(),
                        report_name=report_name,
                        date_range_label=date_range_label,
                        kpi_summary=kpis,
                        waste_summary=waste_summary,
                        match_type_revenue_rows=results.get("match_type_revenue_rows", []),
                        match_type_inefficient_rows=results.get("match_type_inefficient_rows", []),
                        campaign_rows=normalize_records_for_sheet(campaign_summary),
                        campaign_type_rows=results.get("campaign_type_rows", []),
                        top_keyword_rows=simplify_raw_for_sheet(keyword_spend_table, "target", limit=20),
                        top_search_term_rows=simplify_raw_for_sheet(search_term_spend_table, "customer_search_term", limit=20),
                        waste_keyword_rows=simplify_raw_for_sheet(waste_kw_raw, "target", limit=50),
                        waste_search_term_rows=simplify_raw_for_sheet(waste_st_raw, "customer_search_term", limit=50),
                        winner_keyword_rows=simplify_raw_for_sheet(kw_winners_raw, "target", limit=50),
                        winner_search_term_rows=simplify_raw_for_sheet(st_winners_raw, "customer_search_term", limit=50),
                        targeting_data_rows=normalize_records_for_sheet(safe_df(results.get("targeting_with_share"))),
                        search_term_data_rows=normalize_records_for_sheet(safe_df(results.get("search_terms"))),
                    )

                    send_lead_webhook(
                        {
                            "name": lead_name.strip(),
                            "email": lead_email.strip(),
                            "phone": lead_phone.strip(),
                            "brand_name": lead_brand.strip(),
                            "audit_url": created_report.get("url", ""),
                            "report_name": report_name,
                            "date_range_label": date_range_label,
                            "source": "sales_audit_tool",
                        }
                    )

                    st.session_state["sales_audit_created_report"] = created_report
                    st.session_state["sales_audit_unlocked"] = True
                    st.rerun()
                except Exception as exc:
                    st.error(f"We hit an issue while creating the audit: {exc}")

if results and st.session_state.get("sales_audit_unlocked", False):
    created_report = safe_dict(st.session_state.get("sales_audit_created_report", {}))
    kpis = safe_dict(results.get("kpi_summary"))
    keyword_spend_table = safe_df(results.get("keyword_spend_table"))
    waste_tables = safe_dict(results.get("waste_tables"))

    waste_kw_raw = pd.concat(
        [safe_df(waste_tables.get("keyword_zero_sale")), safe_df(waste_tables.get("keyword_high_acos"))],
        ignore_index=True,
    ).drop_duplicates()
    waste_st_raw = pd.concat(
        [safe_df(waste_tables.get("search_zero_sale")), safe_df(waste_tables.get("search_high_acos"))],
        ignore_index=True,
    ).drop_duplicates()
    biggest_waste = waste_st_raw if not waste_st_raw.empty else waste_kw_raw
    waste_term_col = "customer_search_term" if not waste_st_raw.empty else "target"

    st.success("Your branded audit is ready.")
    if created_report.get("url"):
        st.markdown(f"[Open Google Sheet audit]({created_report['url']})")

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f'<div class="metric"><div class="metric-label">Spend</div><div class="metric-value">{format_currency(kpis.get("spend"))}</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="metric"><div class="metric-label">Ad Sales</div><div class="metric-value">{format_currency(kpis.get("ad_sales"))}</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="metric"><div class="metric-label">TACOS</div><div class="metric-value">{format_percent(kpis.get("tacos_pct"))}</div></div>', unsafe_allow_html=True)

    d1, d2 = st.columns(2)
    with d1:
        st.markdown("### Top Keywords")
        top_kw = simplify_term_table(keyword_spend_table, "target", limit=10)
        if not top_kw.empty:
            st.dataframe(top_kw, use_container_width=True)
        else:
            st.info("No keyword data available.")
    with d2:
        st.markdown("### Biggest Waste")
        waste_view = simplify_term_table(biggest_waste, waste_term_col, limit=10)
        if not waste_view.empty:
            st.dataframe(waste_view, use_container_width=True)
        else:
            st.info("No major waste terms found.")
