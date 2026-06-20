/******************************************************************************
 * USDC Wellbeing Hub — Outlook Email Relay
 *
 * A tiny, single-purpose Express service: it receives a "send this email"
 * request from your Google Apps Script backend and relays it through your
 * gopi@usdcglobal.com Microsoft 365 / Outlook account via SMTP, instead of
 * Gmail's MailApp (which has a very low daily sending limit on a personal
 * Gmail account).
 *
 * Apps Script never sees your Outlook password — it only knows a shared
 * "relay key" that this service checks before sending anything.
 ******************************************************************************/

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const nodemailer = require('nodemailer');

const app = express();
app.use(cors());
app.use(express.json({ limit: '20mb' })); // generous limit so PDF attachments fit

const PORT = process.env.PORT || 3000;

// ---- Required environment variables (set these in Render → Environment) ----
const EMAIL_USER = process.env.EMAIL_USER;     // e.g. gopi@usdcglobal.com
const EMAIL_PASS = process.env.EMAIL_PASS;     // Outlook/Microsoft 365 password or app password
const RELAY_KEY = process.env.RELAY_KEY;       // shared secret — must match EMAIL_RELAY_KEY in Code.gs
const EMAIL_FROM_NAME = process.env.EMAIL_FROM_NAME || 'The Wellbeing Hub';
const SMTP_HOST = process.env.SMTP_HOST || 'smtp.office365.com';
const SMTP_PORT = Number(process.env.SMTP_PORT || 587);

if (!EMAIL_USER || !EMAIL_PASS || !RELAY_KEY) {
  console.warn('⚠️  Missing one or more required environment variables: EMAIL_USER, EMAIL_PASS, RELAY_KEY');
}

const transporter = nodemailer.createTransport({
  host: SMTP_HOST,
  port: SMTP_PORT,
  secure: false, // STARTTLS on port 587
  auth: { user: EMAIL_USER, pass: EMAIL_PASS },
  tls: { ciphers: 'TLSv1.2' }
});

// Health check — visiting the bare URL in a browser should show this.
app.get('/', (req, res) => {
  res.send('USDC Wellbeing Hub email relay is running. POST to /api/send-mail to send an email.');
});

/**
 * POST /api/send-mail
 * Body: {
 *   key: "shared relay key",
 *   to: "person@example.com",
 *   cc: "optional, comma-separated",
 *   subject: "Email subject",
 *   html: "<p>HTML body</p>",
 *   attachments: [{ filename, contentBase64, contentType }]   // optional
 * }
 */
app.post('/api/send-mail', async (req, res) => {
  try {
    const { key, to, cc, subject, html, attachments } = req.body || {};

    if (!key || key !== RELAY_KEY) {
      return res.status(401).json({ status: 'error', message: 'Unauthorized — relay key mismatch.' });
    }
    if (!to || !subject || !html) {
      return res.status(400).json({ status: 'error', message: 'Missing required field: to, subject, or html.' });
    }

    const mailAttachments = (attachments || []).map(a => ({
      filename: a.filename || 'attachment.pdf',
      content: Buffer.from(a.contentBase64, 'base64'),
      contentType: a.contentType || 'application/pdf'
    }));

    await transporter.sendMail({
      from: `"${EMAIL_FROM_NAME}" <${EMAIL_USER}>`,
      to,
      cc: cc || undefined,
      subject,
      html,
      attachments: mailAttachments
    });

    res.json({ status: 'ok' });
  } catch (err) {
    console.error('send-mail error:', err);
    res.status(500).json({ status: 'error', message: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Email relay listening on port ${PORT}`);
});
