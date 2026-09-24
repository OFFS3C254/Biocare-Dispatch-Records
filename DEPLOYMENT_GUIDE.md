# Biocare Health Systems — Daily Dispatch Portal
## Vercel Cloud Deployment & Multi-User Setup Guide

This web portal connects directly to your live **Biocare Dispatch Tracker Google Sheet**, allowing multiple team members to manage dispatches simultaneously from any smartphone, tablet, or computer.

---

### Project Structure for Vercel

```
Dispatch/
├── Index.html          # Modern, responsive web app (Minimalist Modern with Electric Blue & Slate)
├── code.gs             # Google Apps Script Web App, executive PDF generator & team email dispatcher
├── vercel.json         # Vercel deployment routes and serverless rewrites
├── api/
│   └── proxy.js        # Vercel Serverless Function proxying to Google Apps Script with CORS
├── images/             # Biocare Health Systems brand assets (biocare-logo.png)
└── DEPLOYMENT_GUIDE.md # This deployment guide
```

---

## 2-Step Deployment to Vercel

### Step 1: Deploy `code.gs` in your Google Sheet (5 minutes)

1. Open your **Biocare Dispatch Tracker** in Google Sheets.
2. In the top menu, click **Extensions > Apps Script**.
3. If there is existing code in `Code.gs`, delete it and paste the entire content of [`code.gs`](file:///c:/Users/Administrator/Desktop/Dispatch/code.gs).
4. Click the blue **Deploy** button (top right) > **New deployment**.
5. Click the gear icon next to "Select type" and choose **Web app**.
6. Fill in the deployment details:
   - **Description**: `Biocare Dispatch Webhook API`
   - **Execute as**: **Me** (`your-email@gmail.com`)
   - **Who has access**: **Anyone** *(Essential: allows your team and the Vercel portal to read/write)*
7. Click **Deploy**.
8. If prompted, click **Authorize access**, select your Google account, click *Advanced*, then *Go to Biocare Dispatch Tracker (unsafe)*, and click *Allow*.
9. **Copy the Web App URL** displayed under *Web app*:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```

---

### Step 2: Deploy to Vercel (2 minutes)

#### Method A: Using GitHub & Vercel Dashboard (Easiest)
1. Push this `Dispatch` directory to your GitHub repository (e.g. `biocare-dispatch`).
2. Go to [vercel.com](https://vercel.com) and log in.
3. Click **Add New... > Project** and import your repository.
4. Under **Environment Variables**, add:
   - **Key**: `APPS_SCRIPT_URL`
   - **Value**: *(Paste the Google Apps Script Web App URL from Step 1)*
5. Click **Deploy**.
6. In ~30 seconds, Vercel will give you a permanent live URL (e.g. `https://biocare-dispatch.vercel.app`)!

#### Method B: Using Vercel CLI
From your terminal in this directory:
```bash
cd C:\Users\Administrator\Desktop\Dispatch
npx vercel
```
Follow the quick prompts:
- Set up and deploy: **y**
- Which scope: *(select your account)*
- Link to existing project: **n**
- Project name: `biocare-dispatch`
- Directory: `./`
- Modify settings: **n**

Then add the environment variable:
```bash
npx vercel env add APPS_SCRIPT_URL production
```
Paste your Google Apps Script URL, and run:
```bash
npx vercel --prod
```

---

### How Your Team Uses the Portal

1. Share the Vercel link (`https://your-project.vercel.app`) with your operations, warehouse, and sales managers.
2. Team members can bookmark it on their smartphones or laptops.
3. Every order added, status changed, or batch imported immediately updates your **Google Sheet** in real time!
