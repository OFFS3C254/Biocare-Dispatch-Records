# Biocare-Dispatch-Records
### Biocare Health Systems Ltd. — Daily Dispatch Portal & Supabase PostgreSQL Database

A cloud-based real-time dispatch management system built for **Biocare Health Systems Ltd.** Powered by **Supabase (PostgreSQL)**, it enables team members across logistics, warehouse, and sales departments to manage daily order dispatches simultaneously with zero latency, generate executive PDF manifests, and track operational metrics.

---

## Key Features

- **Minimalist Modern Design System**: Built with an executive aesthetic featuring Electric Blue gradient accents (`#0052FF` → `#4D7CFF`), Slate-900 contrast, and dual-font typography pairing Calistoga, Inter, and JetBrains Mono.
- **Supabase PostgreSQL Cloud Backend**: Direct, low-latency database connectivity with sub-100ms response times, atomic CRUD operations, and full multi-user concurrency across mobile and desktop.
- **Biocare Brand Integration**: Authentic high-resolution Biocare Health Systems Ltd. logo embedded for web view and generated dispatch manifests.
- **Automated Direct Background Email & PDF Dispatch**: Server SMTP relay via free Gmail App Password (attaches full publication-grade PDF reports) with zero third-party plan fees, plus instant 1-click browser PDF download and mail client composer fallback. Sends shift reports directly from `muithyaalex2@gmail.com` to `alexandremuithya@gmail.com` and `biocarehealthsystems@gmail.com`.
- **Executive PDF Manifests with Simulated Digital Signatures**: One-click generation of publication-grade PDF shift manifests with KPI metrics, delivery distribution ratio bars, courier breakdown chips, itemized status badges, and official simulated digital signatures:
  - **Prepared By (Dispatch Operations)**: Alex Mulwa (*italicized & strikethrough*, stamped with the dynamic current date).
  - **Verified By (Warehouse Logistics Lead)**: Wycliffe Adamba (*strikethrough*, stamped with the active shift's operational date).
- **Batch Order Importer**: Supports drag-and-drop Excel files (`.xlsx`), ERP spreadsheets, or direct clipboard copy-paste with automatic order deduplication and slot validation.
- **100% Turnkey Cloud Deployment**: Fully deployable to Vercel or GitHub Pages in under two minutes with zero server configuration.

---

## Project Structure

```
Biocare-Dispatch-Records/
├── Index.html                    # Responsive Single-Page Application
├── schema.sql                    # PostgreSQL table definitions, RLS policies & 707 records seed
├── biocare_dispatches_master.csv   # Master CSV export of 707 historical Excel orders
├── .env                          # Supabase configuration credentials
├── .env.example                  # Supabase environment template
├── vercel.json                   # Vercel deployment routes and SPA rewrites
├── images/                       # Official Biocare brand assets (biocare-logo.png)
├── DEPLOYMENT_GUIDE.md           # Step-by-step Supabase & Vercel deployment guide
└── README.md                     # Project documentation
```

---

## Quick Start & Deployment

1. **Initialize Database**: In your [Supabase Dashboard](https://supabase.com/dashboard/project/klykpcakybreybrrpbhp) SQL Editor, run [`schema.sql`](schema.sql) to create the `dispatches` table and seed all 707 historical orders.
2. **Deploy to Vercel**: Import this GitHub repository into Vercel and click **Deploy**.

For detailed setup instructions, refer to [`DEPLOYMENT_GUIDE.md`](DEPLOYMENT_GUIDE.md).
