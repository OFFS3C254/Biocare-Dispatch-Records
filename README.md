# Biocare-Dispatch-Records
### Biocare Health Systems Ltd. — Daily Dispatch Portal & Database

A cloud-based daily dispatch management portal built for **Biocare Health Systems Ltd.** It connects directly to the operational Google Sheet (`Biocare Dispatch Tracker`), enabling multi-user concurrency across mobile and desktop devices, automated executive PDF manifest generation, and direct team email dispatch.

---

## Key Features

- **Minimalist Modern Design System**: Built with a curated aesthetic featuring Electric Blue gradient accents (`#0052FF` → `#4D7CFF`), Slate-900 contrast, and typography paired with Calistoga, Inter, and JetBrains Mono.
- **Biocare Brand Integration**: Authentic high-resolution Biocare Health Systems logo embedded for both web view and generated PDF manifests.
- **Automated Executive PDF Reports**: One-click generation of publication-grade PDF shift manifests with KPI metrics, delivery distribution ratio bars, courier breakdown chips, itemized status badges, and official warehouse sign-off lines.
- **Team Email Dispatch**: Directly emails the generated PDF along with a responsive executive HTML shift summary to configured team recipient emails (`REPORT_RECIPIENT_EMAILS`).
- **Batch Order Importer**: Supports drag-and-drop Excel files (`.xlsx`), ERP spreadsheets, or direct clipboard copy-paste with automatic order deduplication and slot validation.
- **Multi-Device Cloud Access**: Deployed on Vercel with zero server maintenance, enabling drivers, warehouse staff, and management to collaborate simultaneously.

---

## Project Structure

```
Dispatch/
├── Index.html          # Frontend web app (Minimalist Modern UI)
├── code.gs             # Google Apps Script Web App & REST API backend
├── vercel.json         # Vercel routing and serverless rewrites
├── api/
│   └── proxy.js        # Vercel Serverless Function proxying to Google Apps Script
├── images/             # Biocare brand assets (biocare-logo.png)
├── DEPLOYMENT_GUIDE.md # Detailed deployment instructions
└── README.md           # Project documentation
```

---

## Deployment & Setup

Refer to [`DEPLOYMENT_GUIDE.md`](DEPLOYMENT_GUIDE.md) for full step-by-step instructions:

1. **Google Apps Script**: Paste [`code.gs`](code.gs) into your Google Sheet's Apps Script editor and deploy as a Web App (access: *Anyone*).
2. **Vercel**: Import this GitHub repository into Vercel and set `APPS_SCRIPT_URL` to your deployed Apps Script URL.
