# Biocare Health Systems Ltd. — Dispatch Operations Portal
## Supabase PostgreSQL Cloud Database & Vercel Deployment Guide

This portal is a real-time dispatch management system built for **Biocare Health Systems Ltd.** Powered by **Supabase (PostgreSQL)**, it enables team members across warehouse, logistics, and sales departments to manage orders simultaneously from any smartphone, tablet, or desktop with sub-100ms latency.

---

### Project Architecture

```
Biocare-Dispatch-Records/
├── Index.html                  # Responsive Single-Page Application (Minimalist Modern UI)
├── schema.sql                  # PostgreSQL table definitions, RLS security policies & 707 records seed
├── biocare_dispatches_master.csv # Master CSV export of all 707 historical orders
├── .env                        # Supabase credentials (URL & Publishable Key)
├── .env.example                # Supabase environment template
├── vercel.json                 # Vercel deployment routes and SPA rewrites
├── images/                     # Official Biocare brand assets (biocare-logo.png)
├── DEPLOYMENT_GUIDE.md         # This deployment guide
└── README.md                   # Project overview & documentation
```

---

## 2-Step Cloud Deployment

### Step 1: Initialize the Supabase Database (1 minute)

Your Supabase project is already configured:
- **Project URL**: `https://klykpcakybreybrrpbhp.supabase.co`
- **Publishable Key**: `sb_publishable_GtvJBDlVpZwLaKNCrjqTbg_b_msekq6`

#### Option A: 1-Click SQL Setup (Recommended)
1. Go to your [Supabase Dashboard](https://supabase.com/dashboard/project/klykpcakybreybrrpbhp).
2. Click **SQL Editor** in the left sidebar.
3. Click **New Query**.
4. Copy the entire contents of [`schema.sql`](schema.sql) and paste it into the editor.
5. Click **Run** (or press `Ctrl + Enter`).
6. Done! This immediately:
   - Creates the `public.dispatches` table with indexes for `dispatch_date` and `order_no`.
   - Configures Row Level Security (RLS) policies allowing secure client reads, inserts, updates, and deletes.
   - Creates the `public.settings` table for courier options.
   - Seeds all **707 historical orders** from `Biocare Dispatch Tracker.xlsx` into PostgreSQL!

#### Option B: Spreadsheet Import via Supabase Studio
1. In the Supabase Dashboard, click **Table Editor**.
2. Click **Import data via spreadsheet**.
3. Drag and drop [`biocare_dispatches_master.csv`](biocare_dispatches_master.csv).
4. Set table name to `dispatches` and click **Import Data**.

---

### Step 2: Deploy to Vercel (2 minutes)

Because the credentials are preconfigured directly in `.env` and `Index.html`, deployment is 100% turnkey.

#### Method A: Via Vercel Dashboard (Easiest)
1. Go to [vercel.com](https://vercel.com) and log in.
2. Click **Add New... > Project**.
3. Import your GitHub repository:
   ```
   https://github.com/OFFS3C254/Biocare-Dispatch-Records.git
   ```
4. Framework Preset: **Other** (Root Directory: `./`).
5. Click **Deploy**.
6. In ~20 seconds, your site is live with a permanent URL (e.g. `https://biocare-dispatch.vercel.app`)!

#### Method B: Via Vercel CLI
From your terminal in this repository:
```bash
npx vercel --prod
```

---

## Operational Features

1. **Zero Cold Starts**: PostgreSQL on Supabase responds in under 100ms, eliminating the 4–6 second delays of legacy webhook services.
2. **Instant Multi-User Sync**: Status changes, new consignments, and batch imports reflect immediately across all connected team devices.
3. **Executive PDF Manifest Generator**: Click **Report & Email Hub** to generate an executive, publication-grade dispatch manifest with authentic Biocare branding, KPI performance cards, courier distribution chips, and warehouse sign-off lines.
4. **Shift Summary Dispatch**: One click prepares a formatted dispatch email with all order metrics ready to send to leadership and sales teams (`biocarehealthsystems@gmail.com, alexandremuithya@gmail.com`).
5. **Batch Excel / ERP Importer**: Drag and drop any daily order spreadsheet or paste copied rows directly to populate the shift.
