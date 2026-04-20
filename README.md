# Evolved Commerce Sales Audit Lead Magnet

This package contains a Cloud Run-ready Streamlit app that:
- accepts Amazon Ads report uploads
- processes a prospect-facing sales audit
- gates results behind a lead form
- creates a branded Google Sheets audit via Google Apps Script
- posts the lead to a Salesforce-compatible webhook

## Files
- `app.py` - main standalone app
- `sales_audit_ingestion.py` - audit engine
- `shared_ingestion_utils.py` - parsing helpers
- `apps_script_webhook.gs` - Google Apps Script used to create the branded audit
- `Dockerfile` - Cloud Run container
- `.env.example` - environment variable template

## Required environment variables
At minimum:
- `GOOGLE_SHEET_WEBHOOK_URL`
- `GOOGLE_SHEETS_TEMPLATE_ID`
- `GOOGLE_DRIVE_FOLDER_ID`
- `LEAD_WEBHOOK_URL`
- `APP_BASE_URL`

Optional but recommended:
- `GOOGLE_SCRIPT_SHARED_SECRET`
- `LEAD_WEBHOOK_BEARER_TOKEN`
- `PLACEHOLDER_LEAD_EMAIL`
- `SITE_SOURCE`

## Local run
```bash
pip install -r requirements.txt
export $(cat .env.example | xargs)
streamlit run app.py
```

## Cloud Run deploy
```bash
gcloud run deploy sales-audit-app \
  --source . \
  --region us-central1 \
  --allow-unauthenticated \
  --set-env-vars GOOGLE_SHEET_WEBHOOK_URL=...,GOOGLE_SHEETS_TEMPLATE_ID=...,GOOGLE_DRIVE_FOLDER_ID=...,LEAD_WEBHOOK_URL=...,APP_BASE_URL=https://audit.yourdomain.com
```

For secrets, prefer Secret Manager and `--set-secrets` instead of plain env vars.

## Apps Script setup
1. Create a Google Apps Script project.
2. Paste in `apps_script_webhook.gs`.
3. Set Script Properties:
   - `SHARED_SECRET`
   - `LEAD_LOG_SHEET_ID` (optional)
4. Deploy as a Web App accessible by anyone with the link.
5. Copy the Web App URL into `GOOGLE_SHEET_WEBHOOK_URL`.

## Launch checklist
- Confirm the Google Sheets template has a `Summary` tab with the expected cell layout.
- Confirm the destination Drive folder is writable by the Apps Script identity.
- Confirm the Salesforce webhook accepts JSON in the shape sent by `build_lead_payload()`.
- Map your custom subdomain to the Cloud Run service.
- Test the full flow with real sample reports before launch.
