const nodemailer = require('nodemailer');
const { createClient } = require('@supabase/supabase-js');

function parsePayload(req) {
  const body = req.body || {};
  const to = Array.isArray(body.to)
    ? body.to.map(v => String(v || '').trim()).filter(Boolean)
    : [];
  const subject = String(body.subject || '').trim();
  const htmlBody = String(body.body || '').trim() || 'Please find attached PDF.';
  const fileName = String(body.filename || 'layout.pdf').replace(/[^\w.\-]/g, '_');
  const attachmentBase64 = String(body.attachmentBase64 || '').trim();
  const contentType = String(body.contentType || 'application/pdf').trim();
  return { to, subject, htmlBody, fileName, attachmentBase64, contentType };
}

function validatePayload(payload) {
  if (!payload.to.length) return 'Recipients are required.';
  if (!payload.subject) return 'Subject is required.';
  if (!payload.attachmentBase64) return 'PDF attachment is required.';
  return '';
}

async function sendViaResend(payload) {
  const resendApiKey = process.env.RESEND_API_KEY || '';
  const fromEmail = process.env.MAIL_FROM || '';
  if (!resendApiKey || !fromEmail) {
    throw new Error('Resend not configured. Set RESEND_API_KEY and MAIL_FROM.');
  }

  const response = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${resendApiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: fromEmail,
      to: payload.to,
      subject: payload.subject,
      html: payload.htmlBody.replace(/\n/g, '<br/>'),
      attachments: [
        {
          filename: payload.fileName,
          content: payload.attachmentBase64,
          type: payload.contentType,
          disposition: 'attachment',
        },
      ],
    }),
  });

  const result = await response.json().catch(() => ({}));
  if (!response.ok) {
    const msg = result?.message || result?.error || 'Resend rejected the request.';
    throw new Error(msg);
  }
  return { providerId: result?.id || null };
}

async function sendViaSmtp(payload, provider) {
  const fromEmail = process.env.MAIL_FROM || '';
  const smtpUser = process.env.SMTP_USER || '';
  const smtpPass = process.env.SMTP_PASS || '';

  if (!fromEmail || !smtpUser || !smtpPass) {
    throw new Error('SMTP not configured. Set MAIL_FROM, SMTP_USER, SMTP_PASS.');
  }

  const host = provider === 'gmail' ? 'smtp.gmail.com' : 'smtp.office365.com';
  const port = provider === 'gmail' ? 465 : 587;
  const secure = provider === 'gmail';

  const transporter = nodemailer.createTransport({
    host,
    port,
    secure,
    auth: {
      user: smtpUser,
      pass: smtpPass,
    },
    connectionTimeout: 15000,
    greetingTimeout: 10000,
    socketTimeout: 20000,
    tls: provider === 'microsoft' ? { ciphers: 'TLSv1.2', rejectUnauthorized: true } : undefined,
  });

  const sendPromise = transporter.sendMail({
    from: fromEmail,
    to: payload.to.join(','),
    subject: payload.subject,
    html: payload.htmlBody.replace(/\n/g, '<br/>'),
    attachments: [
      {
        filename: payload.fileName,
        content: payload.attachmentBase64,
        encoding: 'base64',
        contentType: payload.contentType,
      },
    ],
  });

  const timeoutPromise = new Promise((_, reject) => {
    setTimeout(() => reject(new Error('SMTP send timed out. Check SMTP host/auth/firewall settings.')), 30000);
  });

  const info = await Promise.race([sendPromise, timeoutPromise]);

  return { providerId: info?.messageId || null };
}

async function validatePrintMoreSession(req) {
  const sessionToken = String(req.headers['x-session-token'] || '').trim();
  if (!sessionToken) {
    throw new Error('Unauthorized. Missing session token.');
  }

  const supabaseUrl = process.env.SUPABASE_URL || '';
  const supabaseAnonKey = process.env.SUPABASE_ANON_KEY || '';
  if (!supabaseUrl || !supabaseAnonKey) {
    throw new Error('Server auth not configured. Set SUPABASE_URL and SUPABASE_ANON_KEY.');
  }

  const client = createClient(supabaseUrl, supabaseAnonKey);
  const { error } = await client.rpc('list_printmore_layouts', {
    p_session_token: sessionToken,
  });
  if (error) {
    throw new Error('Unauthorized. Session invalid or expired.');
  }
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, message: 'Method not allowed.' });
  }

  try {
    await validatePrintMoreSession(req);
  } catch (err) {
    return res.status(401).json({ ok: false, message: err?.message || 'Unauthorized.' });
  }

  const payload = parsePayload(req);
  const invalid = validatePayload(payload);
  if (invalid) {
    return res.status(400).json({ ok: false, message: invalid });
  }

  const provider = String(process.env.EMAIL_PROVIDER || 'resend').toLowerCase();
  if (!['resend', 'gmail', 'microsoft'].includes(provider)) {
    return res.status(500).json({
      ok: false,
      message: 'Invalid EMAIL_PROVIDER. Use resend, gmail, or microsoft.',
    });
  }

  try {
    let result;
    if (provider === 'resend') {
      result = await sendViaResend(payload);
    } else {
      result = await sendViaSmtp(payload, provider);
    }

    return res.status(200).json({
      ok: true,
      message: 'Email sent successfully.',
      provider,
      providerId: result?.providerId || null,
    });
  } catch (err) {
    return res.status(500).json({ ok: false, message: err?.message || 'Email send failed.' });
  }
};
