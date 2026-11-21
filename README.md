# AI Forecast Engineer — Today's Working Showcase

> **Status:** runnable MVP demonstration using **100% synthetic financial data**. It is **not** a production system, and it does **not** connect to SAP IBP, Microsoft 365, Salesforce, or private enterprise data.

## One-command local start

Python 3.11+ recommended.

```bash
pip install -r requirements.txt
streamlit run app.py
```

The browser opens a dashboard containing six Excel regional submissions (216 forecast rows) and five real seeded data-quality defects.

## Demo script — ~75 seconds

1. Open **Command Center** and show validated revenue by month and gross-to-net composition.
2. Open **Mapping Studio** and switch between AMER West and EMEA UK to demonstrate differently named/reordered spreadsheet columns mapped into one schema.
3. Open **Exception Investigator**, select **FORMULA_DRIFT**. Show the exact workbook/row and unexpected `+1000` in the source formula.
4. Switch to **Scenario Lab**; move demand to **+12%** and incremental rebates to **+75 bps**. Show the net-revenue impact updating from actual deterministic calculation.
5. Open **AI Analysis**. When `OPENAI_API_KEY` is configured, click **Generate live AI investigation**, then show sourced priorities. Without a key, emphasize the deterministic summary and that live LLM calls are disabled.
6. End with **Export & Method** to download validated-only CSV; explain SAP integration is future work.

## Optional real AI (server-side key only)

Set the API key as an environment variable or in a Streamlit deployment secret:

```
OPENAI_API_KEY = "your-secret-api-key"
```

Do not paste API keys into public posts, upload them to GitHub, or hardcode them in the app. The app calls the OpenAI Responses API **only when the button is pressed**. Model ID is editable (default `gpt-6.2-astra`; use an ID supported by your account). It submits only a concise exception summary of the currently selected workbook rows. It does not pass full workbook contents or let the LLM modify calculations. API charges may apply.

For any public deployment, **disable or gate the live LLM button** if you don't want strangers generating API charges. This app is a showcase, not an abuse-resistant hosted service.

## Deploy today with Streamlit Community Cloud

1. Create a **new GitHub repository** with the files from this folder (including `sample_files/` and `requirements.txt`).
2. Sign in to https://share.streamlit.io/ and choose **Create app** (or **New app**), selecting your repository and `app.py`.
3. Deploy. Share the generated `*.streamlit.app` URL on LinkedIn. No external API key is needed for the dashboard, mapping, anomaly detection, or what-if scenarios.
4. Keep synthetic-data labeling visible. Do not upload confidential spreadsheets to a public deployment.

Optional Replit workflow: import the GitHub repository as a Python project; set run command to `pip install -r requirements.txt && streamlit run app.py --server.address 0.0.0.0 --server.port 8501` and configure the corresponding webview port.

## What works / what does not

| Capability | Status |
|---|---|
| Parse 6 real XLSX files and 216 forecast rows | Works |
| Support 2 header layouts and currency normalization | Works |
| Recalculate revenue with Decimal and synthetic FX | Works |
| Identify 5 planted anomalies with exact row evidence | Works |
| Demand & rebate what-if calculation | Works |
| Optional live LLM-generated investigation | Connect your own key |
| Upload compatible XLSX files | Works; demo-only assumptions |
| Export validated-only canonical CSV | Works |
| Persist approvals, enterprise SSO, multi-tenant RBAC | Not implemented |
| Fully execute Excel formulas and macros | Not implemented (intentionally) |
| SAP IBP writeback | Not implemented |
| Machine-learned or LLM-driven column mapping | Not implemented (approved aliases, deterministic) |

## Quality checks

```bash
pip install pytest
python -m pytest -q tests
```

Test cases assert 216 rows, all five seeded exception categories, full mapping coverage on the fixtures, neutral what-if arithmetic, and USD totals against a separate ground-truth ledger. The ground truth is used **only by tests**, never by the app.

## Architecture for future production implementation

`SharePoint/Excel → isolated parser → deterministic normalization/rules → versioned PostgreSQL → AI investigation (optional) → finance approval → SAP IBP adapter`.

The current Streamlit code is a deliberately compact proof of concept; it is not a substitute for the production reference architecture. Region derivation from filenames and simple additive rebate assumptions are demonstration-only simplifications.
