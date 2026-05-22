const express = require('express');
const cors = require('cors');
const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const app = express();
app.use(cors());
app.use(express.json());

app.post('/save-draft', async (req, res) => {
  try {
    const { tenantId, clientId, clientSecret, outlookEmail, subject, emailBody, toEmail, toName } = req.body;

    // Get token
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: `grant_type=client_credentials&client_id=${clientId}&client_secret=${encodeURIComponent(clientSecret)}&scope=https://graph.microsoft.com/.default`
      }
    );
    const tokenData = await tokenRes.json();
    if (!tokenData.access_token) {
      return res.status(401).json({ error: 'Token failed', detail: tokenData });
    }

    // Save draft
    const draftRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${outlookEmail}/messages`,
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + tokenData.access_token,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          subject,
          body: { contentType: 'Text', content: emailBody },
          toRecipients: toEmail ? [{ emailAddress: { address: toEmail, name: toName || toEmail } }] : [],
          isDraft: true
        })
      }
    );

    const draftData = await draftRes.json();
    if (draftRes.ok) {
      res.json({ success: true });
    } else {
      res.status(400).json({ error: 'Draft failed', detail: draftData });
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/', (req, res) => res.send('TBD Email Proxy — Running'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Proxy running on port ${PORT}`));
