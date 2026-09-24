/**
 * Vercel Serverless Function — Google Apps Script API Proxy (api/proxy.js)
 * ------------------------------------------------------------------------
 * Proxies requests from Vercel to your deployed Google Apps Script Web App.
 * Handles redirects (302) and applies permissive CORS headers.
 */

export default async function handler(req, res) {
  // CORS configuration
  res.setHeader('Access-Control-Allow-Credentials', true);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
  );

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  // Target Google Apps Script Web App URL from environment or request header/param
  const scriptUrl = process.env.APPS_SCRIPT_URL || req.query.scriptUrl || req.headers['x-script-url'];

  if (!scriptUrl) {
    return res.status(400).json({
      success: false,
      error: 'Google Apps Script URL is not configured. Set APPS_SCRIPT_URL environment variable in Vercel settings, or configure it in the portal settings dialog.'
    });
  }

  try {
    let response;
    if (req.method === 'GET') {
      const url = new URL(scriptUrl);
      for (const [key, value] of Object.entries(req.query)) {
        if (key !== 'scriptUrl') {
          url.searchParams.set(key, value);
        }
      }
      response = await fetch(url.toString(), {
        method: 'GET',
        redirect: 'follow'
      });
    } else {
      // POST
      response = await fetch(scriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(req.body),
        redirect: 'follow'
      });
    }

    const text = await response.text();
    try {
      const data = JSON.parse(text);
      return res.status(200).json(data);
    } catch (parseErr) {
      return res.status(200).send(text);
    }
  } catch (err) {
    return res.status(500).json({
      success: false,
      error: 'Proxy request failed: ' + err.message
    });
  }
}
