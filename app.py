import os
from datetime import datetime, timezone
from typing import Any, Optional

import pandas as pd
import requests
import streamlit as st

from sales_audit_ingestion import SalesAuditEngine
from shared_ingestion_utils import to_excel_bytes_multi


st.set_page_config(
    page_title='Amazon Ads Audit | Evolved Commerce',
    page_icon='📈',
    layout='wide',
)

# ------------------------------
# Configuration
# ------------------------------
DEFAULT_HIGH_ACOS_THRESHOLD = float(os.getenv('DEFAULT_HIGH_ACOS_THRESHOLD', '40'))
DEFAULT_WINNING_ACOS_THRESHOLD = float(os.getenv('DEFAULT_WINNING_ACOS_THRESHOLD', '15'))
DEFAULT_MIN_WASTE_SPEND = float(os.getenv('DEFAULT_MIN_WASTE_SPEND', '0.01'))
DEFAULT_MIN_WINNER_ORDERS = int(os.getenv('DEFAULT_MIN_WINNER_ORDERS', '1'))
SHOW_RAW_WORKBOOK_DOWNLOAD = os.getenv('SHOW_RAW_WORKBOOK_DOWNLOAD', 'false').lower() == 'true'
PLACEHOLDER_LEAD_EMAIL = os.getenv('PLACEHOLDER_LEAD_EMAIL', 'jacob@evolvedcommerce.com')
GOOGLE_SHEET_WEBHOOK_URL = os.getenv('GOOGLE_SHEET_WEBHOOK_URL', '')
LEAD_WEBHOOK_URL = os.getenv('LEAD_WEBHOOK_URL', '')
SITE_SOURCE = os.getenv('SITE_SOURCE', 'website_sales_audit')
APP_BASE_URL = os.getenv('APP_BASE_URL', '')
TEMPLATE_ID = os.getenv('GOOGLE_SHEETS_TEMPLATE_ID', '')
DESTINATION_FOLDER_ID = os.getenv('GOOGLE_DRIVE_FOLDER_ID', '')
LEAD_WEBHOOK_BEARER_TOKEN = os.getenv('LEAD_WEBHOOK_BEARER_TOKEN', '')
GOOGLE_SCRIPT_SHARED_SECRET = os.getenv('GOOGLE_SCRIPT_SHARED_SECRET', '')


# ------------------------------
# Styling
# ------------------------------
st.markdown(
    """
    <style>
        .block-container {max-width: 1280px; padding-top: 1.2rem; padding-bottom: 2rem;}
        .hero {
            background: linear-gradient(120deg, #EA580C 0%, #111827 100%);
            border-radius: 24px;
            padding: 38px 34px;
            color: white;
            margin-bottom: 1rem;
            box-shadow: 0 12px 28px rgba(17,24,39,.18);
        }
        .hero h1 {margin: 0 0 .4rem 0; font-size: 2.2rem; line-height: 1.05;}
        .hero p {margin: 0; font-size: 1rem; opacity: .96;}
        .section-title {font-size: 1.2rem; font-weight: 800; color: #111827; margin-top: .8rem;}
        .section-note {font-size: .94rem; color: #6B7280; margin-bottom: .9rem;}
        .metric-card {background: white; border: 1px solid #E5E7EB; border-left: 4px solid #F97316; border-radius: 16px; padding: 14px 16px; min-height: 95px; box-shadow: 0 4px 10px rgba(15, 23, 42, 0.03);}
        .metric-label {font-size: .75rem; text-transform: uppercase; color: #6B7280; font-weight: 700; letter-spacing: .04em; margin-bottom: .25rem;}
        .metric-value {font-size: 1.55rem; font-weight: 800; color: #111827; line-height: 1.1;}
        .status-pill {display: inline-block; padding: 7px 12px; border-radius: 999px; font-size: .84rem; font-weight: 800; margin-bottom: .6rem;}
        .status-good {background: #ECFDF5; color: #047857; border: 1px solid #A7F3D0;}
        .status-warn {background: #FFFBEB; color: #B45309; border: 1px solid #FDE68A;}
        .status-bad {background: #FEF2F2; color: #B91C1C; border: 1px solid #FECACA;}
        .summary-box {background: white; border: 1px solid #E5E7EB; border-radius: 16px; padding: 16px 18px;}
        .step-box {background: #FFF7ED; border: 1px solid #FED7AA; border-radius: 16px; padding: 14px 16px; margin-bottom: .8rem;}
        .small-muted {color: #6B7280; font-size: .9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------
# Helpers
# ------------------------------
def initialize_state() -> None:
    defaults = {
        'audit_results': None,
        'audit_ready': False,
        'lead_submitted': False,
        'created_report': None,
        'lead_webhook_result': None,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def safe_dict(value: Any) -> dict:
    return value if isinstance(value, dict) else {}


def safe_df(value: Any) -> pd.DataFrame:
    return value if isinstance(value, pd.DataFrame) else pd.DataFrame()


def get_number(value: Any, default: float = 0.0) -> float:
    try:
        if value in (None, ''):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def format_currency(value: Any) -> str:
    return f"${get_number(value):,.2f}"


def format_percent(value: Any) -> str:
    return f"{get_number(value):,.2f}%"


def format_number(value: Any) -> str:
    return f"{get_number(value):,.2f}"


def tone_from_health(status: str) -> str:
    lowered = str(status or '').lower().strip()
    if lowered == 'healthy':
        return 'status-good'
    if lowered == 'mixed':
        return 'status-warn'
    return 'status-bad'


def render_metric_card(label: str, value: str) -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def simplify_term_table(df: pd.DataFrame, term_col: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    if 'acos' in out.columns and 'acos_pct' not in out.columns:
        out['acos_pct'] = out['acos'] * 100
    keep_cols = [c for c in [term_col, 'spend', 'sales', 'acos_pct'] if c in out.columns]
    out = out[keep_cols].copy()
    out = out.rename(columns={term_col: 'term', 'acos_pct': 'acos'})
    out['spend'] = pd.to_numeric(out['spend'], errors='coerce').fillna(0)
    out['sales'] = pd.to_numeric(out['sales'], errors='coerce').fillna(0)
    out['acos'] = pd.to_numeric(out['acos'], errors='coerce').fillna(0)
    out['spend'] = out['spend'].map(lambda x: f'${x:,.2f}')
    out['sales'] = out['sales'].map(lambda x: f'${x:,.2f}')
    out['acos'] = out['acos'].map(lambda x: f'{x:,.2f}%')
    return out.reset_index(drop=True)


def simplify_campaign_table(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    keep = [c for c in ['campaign_name', 'spend', 'sales', 'acos_pct', 'campaign_status'] if c in df.columns]
    out = df[keep].copy()
    for col in ['spend', 'sales', 'acos_pct']:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors='coerce').fillna(0).round(2)
    return out.sort_values(['spend', 'sales'], ascending=[False, False]).reset_index(drop=True)


def normalize_records_for_sheet(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        return []
    out = df.copy().replace({pd.NA: None}).where(pd.notnull(df), None)

    def parse_numeric_like(value):
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return value
        text = str(value).strip()
        if not text:
            return None
        negative = text.startswith('(') and text.endswith(')')
        text = text.replace('(', '').replace(')', '').replace('$', '').replace(',', '').replace('%', '').strip()
        try:
            num = float(text)
            return -num if negative else num
        except ValueError:
            return value

    for col in out.columns:
        col_l = str(col).lower().strip()
        out[col] = out[col].map(parse_numeric_like)
        if col_l in {'ctr', 'cvr', 'acos', 'acos_pct'} or 'pct' in col_l or 'percent' in col_l:
            out[col] = pd.to_numeric(out[col], errors='coerce').fillna(0).map(lambda x: round(float(x) / 100.0, 6))
        elif col_l in {'roas'}:
            out[col] = pd.to_numeric(out[col], errors='coerce').fillna(0).map(lambda x: round(float(x), 2))
        elif col_l in {'cpc', 'spend', 'sales', 'ad_sales', 'total_sales', 'organic_sales', 'ntb_sales'}:
            out[col] = pd.to_numeric(out[col], errors='coerce').fillna(0).map(lambda x: round(float(x), 2))
        elif col_l in {'impressions', 'clicks', 'orders', 'units_ordered', 'sessions', 'ntb_orders'}:
            out[col] = pd.to_numeric(out[col], errors='coerce').fillna(0).map(lambda x: int(round(float(x))))
    return out.to_dict('records')


def post_json(url: str, payload: dict, bearer_token: str = '', timeout: int = 180) -> dict:
    headers = {'Content-Type': 'application/json'}
    if bearer_token:
        headers['Authorization'] = f'Bearer {bearer_token}'
    response = requests.post(url, json=payload, headers=headers, timeout=timeout)
    response.raise_for_status()
    if response.content:
        try:
            return response.json()
        except Exception:
            return {'success': True, 'raw_response': response.text}
    return {'success': True}


def build_lead_payload(lead: dict, results: dict, request_meta: dict) -> dict:
    kpis = safe_dict(results.get('kpi_summary'))
    waste = safe_dict(results.get('waste_summary'))
    health = safe_dict(results.get('health_summary'))
    return {
        'source': SITE_SOURCE,
        'submittedAtUtc': datetime.now(timezone.utc).isoformat(),
        'lead': lead,
        'auditSummary': {
            'status': health.get('status', ''),
            'summary': health.get('summary', ''),
            'narrative': results.get('narrative', ''),
            'tacosPct': kpis.get('tacos_pct', 0),
            'acosPct': kpis.get('acos_pct', 0),
            'spend': kpis.get('spend', 0),
            'adSales': kpis.get('ad_sales', 0),
            'totalSales': kpis.get('total_sales', 0),
            'wastedSpend': waste.get('wasted_spend', 0),
            'wastedSpendPct': waste.get('wasted_spend_pct', 0),
            'dateRangeLabel': results.get('date_range_label', ''),
        },
        'requestMeta': request_meta,
    }


def create_google_sheet_report(lead: dict, results: dict) -> dict:
    if not GOOGLE_SHEET_WEBHOOK_URL:
        raise RuntimeError('GOOGLE_SHEET_WEBHOOK_URL is not configured.')
    if not TEMPLATE_ID or not DESTINATION_FOLDER_ID:
        raise RuntimeError('Google Sheet template or destination folder ID is not configured.')

    brand_name = lead['brand_name']
    report_name = f"{brand_name} - Amazon Ads Audit"
    kpis = safe_dict(results.get('kpi_summary'))
    waste_summary = safe_dict(results.get('waste_summary'))
    campaign_summary = safe_df(results.get('campaign_summary'))
    keyword_spend_table = safe_df(results.get('keyword_spend_table'))
    search_term_spend_table = safe_df(results.get('search_term_spend_table'))
    waste_tables = safe_dict(results.get('waste_tables'))
    winner_tables = safe_dict(results.get('winner_tables'))

    top_kw = simplify_term_table(keyword_spend_table, 'target').head(20)
    top_st = simplify_term_table(search_term_spend_table, 'customer_search_term').head(20)
    kw_zero = simplify_term_table(safe_df(waste_tables.get('keyword_zero_sale')), 'target')
    kw_high = simplify_term_table(safe_df(waste_tables.get('keyword_high_acos')), 'target')
    st_zero = simplify_term_table(safe_df(waste_tables.get('search_zero_sale')), 'customer_search_term')
    st_high = simplify_term_table(safe_df(waste_tables.get('search_high_acos')), 'customer_search_term')
    kw_winners = simplify_term_table(safe_df(winner_tables.get('keyword_winners')), 'target').head(20)
    st_winners = simplify_term_table(safe_df(winner_tables.get('search_winners')), 'customer_search_term').head(20)

    payload = {
        'sharedSecret': GOOGLE_SCRIPT_SHARED_SECRET,
        'templateId': TEMPLATE_ID,
        'destinationFolderId': DESTINATION_FOLDER_ID,
        'reportName': report_name,
        'brandName': brand_name,
        'dateRangeLabel': (results.get('date_range_label') or '').strip() or 'MM/DD - MM/DD',
        'kpiSummary': kpis,
        'wasteSummary': waste_summary,
        'matchTypeRevenueRows': results.get('match_type_revenue_rows', []),
        'matchTypeInefficientRows': results.get('match_type_inefficient_rows', []),
        'campaignRows': normalize_records_for_sheet(campaign_summary),
        'campaignTypeRows': results.get('campaign_type_rows', []),
        'topKeywordRows': normalize_records_for_sheet(top_kw),
        'topSearchTermRows': normalize_records_for_sheet(top_st),
        'wasteKeywordRows': normalize_records_for_sheet(pd.concat([kw_zero, kw_high], ignore_index=True).drop_duplicates()),
        'wasteSearchTermRows': normalize_records_for_sheet(pd.concat([st_zero, st_high], ignore_index=True).drop_duplicates()),
        'winnerKeywordRows': normalize_records_for_sheet(kw_winners),
        'winnerSearchTermRows': normalize_records_for_sheet(st_winners),
        'targetingDataRows': normalize_records_for_sheet(safe_df(results.get('targeting_with_share'))),
        'searchTermDataRows': normalize_records_for_sheet(safe_df(results.get('search_terms'))),
        'lead': {
            'name': lead['name'],
            'email': lead['email'],
            'phone': lead['phone'],
            'brandName': brand_name,
            'source': SITE_SOURCE,
        },
    }
    data = post_json(GOOGLE_SHEET_WEBHOOK_URL, payload, timeout=180)
    if not data.get('success', False):
        raise RuntimeError(data.get('error', 'Unknown Apps Script error'))
    return data


def send_lead_webhook(lead: dict, results: dict, report_url: str = '') -> dict:
    if not LEAD_WEBHOOK_URL:
        return {'success': False, 'skipped': True, 'message': 'LEAD_WEBHOOK_URL not configured.'}
    payload = build_lead_payload(
        lead,
        results,
        {
            'appBaseUrl': APP_BASE_URL,
            'googleSheetUrl': report_url,
            'placeholderNotificationEmail': PLACEHOLDER_LEAD_EMAIL,
        },
    )
    return post_json(LEAD_WEBHOOK_URL, payload, bearer_token=LEAD_WEBHOOK_BEARER_TOKEN, timeout=60)


def clear_audit_state():
    st.session_state['audit_results'] = None
    st.session_state['audit_ready'] = False
    st.session_state['lead_submitted'] = False
    st.session_state['created_report'] = None
    st.session_state['lead_webhook_result'] = None


# ------------------------------
# UI
# ------------------------------
initialize_state()

st.markdown(
    """
    <div class="hero">
      <h1>Free Amazon Account Audit | Evolved Commerce</h1>
      <p>Upload your recent Amazon reports and get a personalized audit that highlights wasted spend, top opportunities, and winning terms.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown('<div class="step-box"><strong>1. Upload 6 reports</strong><br><span class="small-muted">Recent reporting windows work best.</span></div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="step-box"><strong>2. Run your audit</strong><br><span class="small-muted">We process the reports and prepare your personalized analysis.</span></div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="step-box"><strong>3. Unlock results</strong><br><span class="small-muted">Submit your details to receive the completed audit.</span></div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown('## What to upload')
    st.markdown('- Bulk Sheet\n- Sponsored Products Search Term Report\n- Sponsored Products Targeting Report\n- Sponsored Products Impression Share Report\n- Sales & Traffic Business Report\n- Sponsored Brands Campaign Report (optional)')
    st.markdown('---')
    st.markdown('## Notes')
    st.markdown('- Keep the date ranges aligned where possible\n- The SB report helps populate new-to-brand metrics\n- Your uploaded data is only used to generate this audit')
    if st.button('Start over', use_container_width=True):
        clear_audit_state()
        st.rerun()

st.markdown('<div class="section-title">Brand & Report Uploads</div>', unsafe_allow_html=True)
st.markdown('<div class="section-note">Enter your brand name, then upload the required files.</div>', unsafe_allow_html=True)

brand_name_input = st.text_input('Brand Name', placeholder='Your brand name', key='brand_name_input')

u1, u2 = st.columns(2)
u3, u4 = st.columns(2)
u5, u6 = st.columns(2)
with u1:
    bulk_file = st.file_uploader('Bulk Sheet', type=['xlsx', 'xls', 'csv'], key='bulk_file')
with u2:
    impression_share_file = st.file_uploader('SP Impression Share Report', type=['xlsx', 'xls', 'csv'], key='impression_share_file')
with u3:
    targeting_file = st.file_uploader('SP Targeting Report', type=['xlsx', 'xls', 'csv'], key='targeting_file')
with u4:
    search_term_file = st.file_uploader('SP Search Term Report', type=['xlsx', 'xls', 'csv'], key='search_term_file')
with u5:
    business_report_file = st.file_uploader('Sales & Traffic Business Report', type=['xlsx', 'xls', 'csv'], key='business_report_file')
with u6:
    sb_campaign_file = st.file_uploader('Sponsored Brands Campaign Report (optional)', type=['xlsx', 'xls', 'csv'], key='sb_campaign_file')

required_ready = all([
    brand_name_input.strip(),
    bulk_file is not None,
    impression_share_file is not None,
    targeting_file is not None,
    search_term_file is not None,
    business_report_file is not None,
])

st.caption('Ready to run' if required_ready else 'Please add your brand name and all 5 required reports.')

if st.button('Audit My Account', type='primary', use_container_width=True, disabled=not required_ready):
    clear_audit_state()
    with st.spinner('Running your audit...'):
        try:
            engine = SalesAuditEngine(
                bulk_file=bulk_file,
                impression_share_file=impression_share_file,
                targeting_file=targeting_file,
                search_term_file=search_term_file,
                business_report_file=business_report_file,
                sb_campaign_file=sb_campaign_file,
                high_acos_threshold=DEFAULT_HIGH_ACOS_THRESHOLD,
                winning_acos_threshold=DEFAULT_WINNING_ACOS_THRESHOLD,
                min_waste_spend=DEFAULT_MIN_WASTE_SPEND,
                min_winner_orders=DEFAULT_MIN_WINNER_ORDERS,
                brand_name=brand_name_input.strip(),
            )
            results = engine.process()
            results['brand_name'] = brand_name_input.strip()
            results['has_sb_report'] = sb_campaign_file is not None
            st.session_state['audit_results'] = results
            st.session_state['audit_ready'] = True
            st.success('Your audit is ready to unlock.')
        except Exception as exc:
            st.error(f'We could not process those files: {exc}')

results = safe_dict(st.session_state.get('audit_results'))

if st.session_state.get('audit_ready') and not st.session_state.get('lead_submitted'):
    st.markdown('<div class="section-title">Unlock Your Branded Audit</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-note">Submit your details and we will generate your branded Google Sheets audit.</div>', unsafe_allow_html=True)

    with st.form('lead_capture_form', clear_on_submit=False):
        l1, l2 = st.columns(2)
        with l1:
            lead_name = st.text_input('Full Name')
            lead_email = st.text_input('Work Email')
        with l2:
            lead_phone = st.text_input('Phone Number')
            lead_brand_name = st.text_input('Brand Name', value=results.get('brand_name', brand_name_input.strip()))
        submitted = st.form_submit_button('Unlock My Audit', type='primary', use_container_width=True)

    if submitted:
        if not lead_name.strip() or not lead_email.strip() or not lead_brand_name.strip():
            st.error('Please complete your name, email, and brand name.')
        else:
            lead = {
                'name': lead_name.strip(),
                'email': lead_email.strip(),
                'phone': lead_phone.strip(),
                'brand_name': lead_brand_name.strip(),
            }
            with st.spinner('Generating your branded audit...'):
                try:
                    created_report = create_google_sheet_report(lead, results)
                    lead_response = send_lead_webhook(lead, results, report_url=created_report.get('url', ''))
                    st.session_state['created_report'] = created_report
                    st.session_state['lead_webhook_result'] = lead_response
                    st.session_state['lead_submitted'] = True
                    st.success('Your audit has been unlocked.')
                    st.rerun()
                except Exception as exc:
                    st.error(f'We hit an issue while creating the audit: {exc}')

if st.session_state.get('lead_submitted'):
    created_report = safe_dict(st.session_state.get('created_report'))
    lead_webhook_result = safe_dict(st.session_state.get('lead_webhook_result'))
    brand_name = str(results.get('brand_name', '')).strip()
    kpis = safe_dict(results.get('kpi_summary'))
    waste_summary = safe_dict(results.get('waste_summary'))
    health_summary = safe_dict(results.get('health_summary'))
    narrative = str(results.get('narrative', '')).strip()
    campaign_summary = safe_df(results.get('campaign_summary'))
    keyword_spend_table = safe_df(results.get('keyword_spend_table'))
    search_term_spend_table = safe_df(results.get('search_term_spend_table'))
    waste_tables = safe_dict(results.get('waste_tables'))
    winner_tables = safe_dict(results.get('winner_tables'))

    top_kw = simplify_term_table(keyword_spend_table, 'target').head(15)
    top_st = simplify_term_table(search_term_spend_table, 'customer_search_term').head(15)
    kw_winners = simplify_term_table(safe_df(winner_tables.get('keyword_winners')), 'target').head(15)
    st_winners = simplify_term_table(safe_df(winner_tables.get('search_winners')), 'customer_search_term').head(15)
    waste_kw = pd.concat([
        simplify_term_table(safe_df(waste_tables.get('keyword_zero_sale')), 'target'),
        simplify_term_table(safe_df(waste_tables.get('keyword_high_acos')), 'target'),
    ], ignore_index=True).drop_duplicates().head(15)
    waste_st = pd.concat([
        simplify_term_table(safe_df(waste_tables.get('search_zero_sale')), 'customer_search_term'),
        simplify_term_table(safe_df(waste_tables.get('search_high_acos')), 'customer_search_term'),
    ], ignore_index=True).drop_duplicates().head(15)

    st.markdown('<div class="section-title">Your Audit Is Ready</div>', unsafe_allow_html=True)
    if created_report.get('url'):
        st.success('Your branded Google Sheet report was created successfully.')
        st.markdown(f"[Open your branded audit]({created_report['url']})")
    if lead_webhook_result and not lead_webhook_result.get('skipped'):
        st.caption('Lead webhook sent successfully.' if lead_webhook_result.get('success', True) else 'Lead webhook response received.')

    tone_class = tone_from_health(health_summary.get('status', ''))
    st.markdown(f'<div class="status-pill {tone_class}">{health_summary.get("status", "Unknown")}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="summary-box"><strong>Summary:</strong> {health_summary.get("summary", "")}<br><br><strong>Narrative:</strong> {narrative}</div>', unsafe_allow_html=True)

    r1 = st.columns(4)
    with r1[0]:
        render_metric_card('Spend', format_currency(kpis.get('spend')))
    with r1[1]:
        render_metric_card('Ad Sales', format_currency(kpis.get('ad_sales')))
    with r1[2]:
        render_metric_card('Total Sales', format_currency(kpis.get('total_sales')))
    with r1[3]:
        render_metric_card('Organic Sales', format_currency(kpis.get('organic_sales')))

    r2 = st.columns(4)
    with r2[0]:
        render_metric_card('ACOS', format_percent(kpis.get('acos_pct')))
    with r2[1]:
        render_metric_card('ROAS', format_number(kpis.get('roas')))
    with r2[2]:
        render_metric_card('TACOS', format_percent(kpis.get('tacos_pct')))
    with r2[3]:
        render_metric_card('Wasted Spend', format_currency(waste_summary.get('wasted_spend')))

    st.markdown('<div class="section-title">Top Spend Drivers</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('**Top Keyword / Target Spenders**')
        st.dataframe(top_kw, use_container_width=True) if not top_kw.empty else st.info('No keyword target spend data available.')
    with c2:
        st.markdown('**Top Customer Search Terms**')
        st.dataframe(top_st, use_container_width=True) if not top_st.empty else st.info('No customer search term spend data available.')

    st.markdown('<div class="section-title">Waste & Winners</div>', unsafe_allow_html=True)
    c3, c4 = st.columns(2)
    with c3:
        st.markdown('**Potential Waste**')
        st.dataframe(waste_st if not waste_st.empty else waste_kw, use_container_width=True) if (not waste_st.empty or not waste_kw.empty) else st.info('No major waste terms found.')
    with c4:
        st.markdown('**Winning Terms**')
        st.dataframe(st_winners if not st_winners.empty else kw_winners, use_container_width=True) if (not st_winners.empty or not kw_winners.empty) else st.info('No winner terms found.')

    with st.expander('Campaign Summary', expanded=False):
        campaign_view = simplify_campaign_table(campaign_summary).head(20)
        st.dataframe(campaign_view, use_container_width=True) if not campaign_view.empty else st.info('No campaign summary available.')

    if SHOW_RAW_WORKBOOK_DOWNLOAD:
        export_sheets = {
            'KPI Summary': pd.DataFrame([kpis]),
            'Waste Summary': pd.DataFrame([waste_summary]),
            'Campaign Summary': simplify_campaign_table(campaign_summary),
            'Keyword Spend': simplify_term_table(keyword_spend_table, 'target'),
            'Search Term Spend': simplify_term_table(search_term_spend_table, 'customer_search_term'),
        }
        export_bytes = to_excel_bytes_multi(export_sheets)
        st.download_button(
            label='Download Audit Workbook',
            data=export_bytes,
            file_name=f"{brand_name or 'sales_audit'}_audit_workbook.xlsx".replace(' ', '_'),
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
        )

if not st.session_state.get('audit_ready'):
    st.info('Upload your reports and click Audit My Account to begin.')
