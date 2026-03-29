/* Outlook Email Evaluator - Desktop Add-in (taskpane.js) v2.1 */

Office.onReady(() => {
  initUI();
  loadEmail();
  // Re-load email subject when user switches to a different email (pinned taskpane)
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      loadEmail();
      // Reset UI to ready state for new email
      const btn = document.getElementById('analyze-btn');
      if (btn) { btn.disabled = false; btn.textContent = '🔍 Analyze Email'; }
      document.getElementById('result-body').innerHTML = '<p>Click <strong>Analyze Email</strong> to check this email for threats.</p>';
    });
  } catch(e) {}
});

function storageGet(key) {
  try { return Office.context.roamingSettings.get(key) || ''; }
  catch(e) { try { return localStorage.getItem(key) || ''; } catch(e2) { return ''; } }
}
function storageSet(key, value) {
  try { Office.context.roamingSettings.set(key, value); Office.context.roamingSettings.saveAsync(); }
  catch(e) { try { localStorage.setItem(key, value); } catch(e2) {} }
}

function escapeHtml(s) {
  if (s == null || s === '') return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// --- Gift Card Fraud Detection ---
const GIFT_CARD_KEYWORDS = [
  'gift card', 'gift cards', 'itunes card', 'google play card', 'amazon gift card',
  'steam card', 'ebay gift card', 'visa gift card', 'buy gift cards', 'purchase gift cards',
  'get gift cards', 'send gift cards', 'gift card number', 'gift card code',
  'scratch the card', 'scratch card', 'card balance', 'redeem the card',
  'send me the codes', 'send the codes', 'send the numbers'
];
function checkForGiftCardFraud(subject, body) {
  const combined = ((subject || '') + ' ' + (body || '')).toLowerCase();
  return GIFT_CARD_KEYWORDS.some(kw => combined.includes(kw));
}

function initUI() {
  document.getElementById('settings-btn').addEventListener('click', () => {
    const panel = document.getElementById('settings-panel');
    const main  = document.getElementById('main-panel');
    const open  = panel.classList.toggle('hidden');
    main.classList.toggle('hidden', !open);
    if (!open) populateSettings();
  });
  document.getElementById('save-settings-btn').addEventListener('click', saveSettings);
  document.getElementById('analyze-btn').addEventListener('click', analyzeEmail);
  // Event delegation for finding card toggles
  document.getElementById('result-body').addEventListener('click', (e) => {
    const header = e.target.closest('.oe-finding-header');
    if (header) header.parentElement.classList.toggle('oe-finding-open');
  });
  populateSettings();
}

const DEFAULT_PROXY_URL = 'https://pikplhvawbhndijpkdbq.supabase.co/functions/v1/analyze-email'

function populateSettings() {
  document.getElementById('proxy-url-input').value    = storageGet('proxyUrl') || DEFAULT_PROXY_URL;
  document.getElementById('tenant-domain-input').value = storageGet('tenantDomain');
  document.getElementById('custom-prompt-input').value = storageGet('customPrompt');
}

function saveSettings() {
  storageSet('proxyUrl',      document.getElementById('proxy-url-input').value.trim());
  storageSet('extToken',      document.getElementById('ext-token-input').value.trim());
  storageSet('tenantDomain',  document.getElementById('tenant-domain-input').value.trim());
  storageSet('customPrompt',  document.getElementById('custom-prompt-input').value.trim());
  const msg = document.getElementById('settings-msg');
  msg.textContent = 'Saved!';
  msg.classList.remove('hidden');
  setTimeout(() => msg.classList.add('hidden'), 2000);
}

function loadEmail() {
  const item = Office.context.mailbox.item;
  if (!item) return;
  const subject = item.subject || '(No subject)';
  document.getElementById('email-subject').textContent =
    subject.length > 70 ? subject.slice(0, 70) + '...' : subject;
}

const HIGH_RISK_EXT  = ['.htm','.html','.js','.vbs','.vbe','.ps1','.wsf','.wsh','.jar','.hta'];
const SUSPICIOUS_EXT = ['.exe','.msi','.bat','.cmd','.iso','.img','.zip','.rar','.7z','.docm','.xlsm','.pptm','.lnk'];

function classifyAttachments(attachments) {
  const names      = attachments.map(a => (a.name || '').toLowerCase());
  const highRisk   = names.filter(n => HIGH_RISK_EXT.some(e => n.endsWith(e)));
  const suspicious = names.filter(n => !highRisk.includes(n) && SUSPICIOUS_EXT.some(e => n.endsWith(e)));
  return { names, highRisk, suspicious };
}

// --- SafeLinks / URL-wrapper decoder ---
function decodeWrappedUrl(href) {
  if (!href) return href;
  try {
    href = href.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
    if (href.includes('safelinks.protection.outlook.com')) {
      const u = new URL(href);
      const decoded = u.searchParams.get('url');
      if (decoded) return decodeURIComponent(decoded);
    }
    if (href.includes('trendmicro') || href.includes('imsva') || href.includes('tmase')) {
      const u = new URL(href);
      const decoded = u.searchParams.get('url') || u.searchParams.get('u') || u.searchParams.get('__u');
      if (decoded) return decodeURIComponent(decoded);
      const b64 = u.searchParams.get('redirectUrl') || u.searchParams.get('r');
      if (b64) { try { return atob(b64); } catch(e) {} }
    }
    if (href.includes('urldefense') && href.includes('/v2/')) {
      const u = new URL(href);
      let raw = u.searchParams.get('u');
      if (raw) { raw = raw.replace(/-/g, '%').replace(/_/g, '/'); return decodeURIComponent(raw); }
    }
    if (href.includes('urldefense') && href.includes('/v3/')) {
      const match = href.match(/\/v3\/__([^_]+)__/);
      if (match) return decodeURIComponent(match[1]);
    }
    if (href.includes('mimecast.com')) {
      const u = new URL(href);
      const decoded = u.searchParams.get('url') || u.searchParams.get('u');
      if (decoded) return decodeURIComponent(decoded);
    }
    if (href.includes('?')) {
      const u = new URL(href);
      const decoded = u.searchParams.get('url') || u.searchParams.get('u');
      if (decoded && (decoded.startsWith('http') || decoded.startsWith('%68%74'))) {
        return decodeURIComponent(decoded);
      }
    }
  } catch(e) {}
  return href;
}

function extractLinks(html) {
  if (!html) return [];
  const doc = (new DOMParser()).parseFromString(html, 'text/html');
  const seen = new Set(); const links = [];
  doc.querySelectorAll('a[href]').forEach(a => {
    try {
      const displayText = (a.textContent || '').trim();
      let href = a.getAttribute('href') || '';
      href = decodeWrappedUrl(href);
      if (!href || href.startsWith('mailto:') || href.startsWith('#') || href.length < 10) return;
      let hrefDomain = ''; try { hrefDomain = new URL(href).hostname.toLowerCase(); } catch(e) { return; }
      if (seen.has(hrefDomain)) return; seen.add(hrefDomain);
      let displayDomain = '';
      const urlPattern = displayText.match(/(?:https?:\/\/|www\.)([\w.-]+)/i);
      if (urlPattern) {
        try { displayDomain = new URL(displayText.startsWith('http') ? displayText : 'https://' + displayText).hostname.toLowerCase(); }
        catch(e) { displayDomain = urlPattern[1].toLowerCase(); }
      }
      const mismatch = displayDomain && hrefDomain &&
        !hrefDomain.includes(displayDomain.replace(/^www\./, '')) &&
        !displayDomain.includes(hrefDomain.replace(/^www\./, ''));
      links.push({ display: displayText.slice(0, 80) || '(no text)', href: hrefDomain, fullUrl: href, mismatch });
    } catch(e) {}
  });
  return links.slice(0, 20);
}

async function analyzeEmail() {
  const proxyUrl = storageGet('proxyUrl');
  const extToken = storageGet('extToken');
  if (!proxyUrl) { showError('No proxy URL set. Click the gear icon to configure.'); return; }
  if (!extToken) { showError('No extension token set. Click the gear icon to configure.'); return; }

  setLoading();
  const item = Office.context.mailbox.item;

  const bodyHtml = await new Promise(resolve =>
    item.body.getAsync(Office.CoercionType.Html, r =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '')
    )
  );
  const bodyText = await new Promise(resolve =>
    item.body.getAsync(Office.CoercionType.Text, r =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '')
    )
  );

  const sender      = item.from ? (item.from.displayName + ' <' + item.from.emailAddress + '>') : '(Unknown sender)';
  const subject     = item.subject || '(No subject)';
  const links       = extractLinks(bodyHtml);
  const attachments = item.attachments || [];
  const { names: attachNames, highRisk, suspicious } = classifyAttachments(attachments);
  const tenantDomain = storageGet('tenantDomain') || '';
  const customPrompt = storageGet('customPrompt') || '';

  let isOutlookExternal = false;
  try {
    const senderEmail = item.from ? item.from.emailAddress.toLowerCase() : '';
    if (tenantDomain && senderEmail && !senderEmail.endsWith('@' + tenantDomain.toLowerCase())) {
      isOutlookExternal = true;
    }
    if (Office.context.requirements && Office.context.requirements.isSetSupported('Mailbox', '1.8')) {
      const headers = await new Promise(resolve =>
        item.getAllInternetHeadersAsync(r =>
          resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '')
        )
      );
      if (headers && headers.toLowerCase().includes('x-ms-exchange-organization-scl')) {
        isOutlookExternal = true;
      }
    }
  } catch(e) {}

  if (checkForGiftCardFraud(subject, bodyText)) {
    showResult({
      verdict: 'PHISHING', phishing_score: 99, spam_score: 10,
      summary: 'This email contains a request for gift cards. This is one of the most common fraud tactics used against businesses — it is almost certainly a scam.',
      findings: [{ flag: 'Gift card request detected', explanation: 'Fraudsters impersonate managers, executives, or colleagues and ask employees to buy gift cards urgently. No legitimate business request will ever ask for gift card payments.', howToSpotIt: 'If ANY email asks you to buy gift cards and send the codes — stop immediately. Call that person directly on a known phone number to verify.' }],
      lesson: 'No legitimate business transaction is ever completed with gift cards. If someone asks you to buy gift cards and send the codes, it is a scam — 100% of the time.',
      suggested_action: 'Do NOT purchase any gift cards. Report this email to your IT security team and your manager immediately.'
    }, { subject, links, highRisk, suspicious });
    return;
  }

  const emailData = {
    subject, sender, senderHasEmail: sender.includes('@'),
    body: bodyText.slice(0, 3000), links,
    attachments: attachNames, hasHighRiskAttachment: highRisk.length > 0,
    hasSuspiciousAttachment: suspicious.length > 0, highRiskFiles: highRisk,
    suspiciousFiles: suspicious, isOutlookExternal, clientTimestamp: new Date().toISOString(),
    clientTimezone: Intl.DateTimeFormat().resolvedOptions().timeZone
  };

  try {
    const response = await fetch(proxyUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token: extToken, emailData, customPrompt, tenantDomain })
    });
    if (response.status === 429) { showError('Please wait 5 seconds before analyzing another email.'); return; }
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      showError('Proxy error ' + response.status + ': ' + (err.error || response.statusText));
      return;
    }
    const data = await response.json();
    window._oe_lastResult = data.result;
    showResult(data.result, { subject, links, highRisk, suspicious });
  } catch (err) { showError('Request failed: ' + err.message); }
}

function setLoading() {
  document.getElementById('result-body').innerHTML = '<div class="loading"><div class="spinner"></div><span>Analyzing email...</span></div>';
  document.getElementById('analyze-btn').disabled = true;
  document.getElementById('analyze-btn').textContent = 'Analyzing...';
}

function showError(msg) {
  document.getElementById('result-body').innerHTML = '<div class="error">⚠️ ' + escapeHtml(msg) + '</div>';
  resetBtn();
}

function resetBtn() {
  const btn = document.getElementById('analyze-btn');
  btn.disabled = false;
  btn.textContent = '🔍 Analyze Email';
}

function showResult(result, { subject, links, highRisk, suspicious }) {
  window._oe_lastResult = result;
  const vc = { SAFE:'verdict-safe', SUSPICIOUS:'verdict-suspicious', SPAM:'verdict-spam', PHISHING:'verdict-phishing' }[result.verdict] || 'verdict-suspicious';
  const vi = { SAFE:'✅', SUSPICIOUS:'⚠️', SPAM:'🚫', PHISHING:'🎣' }[result.verdict] || '⚠️';

  const findingsHTML = (result.findings || []).map(f => `
    <div class="oe-finding">
      <div class="oe-finding-header">
        <span class="finding-icon">🚩</span>
        <span class="finding-flag">${escapeHtml(f.flag)}</span>
        <span class="finding-toggle">▼</span>
      </div>
      <div class="oe-finding-body">
        <div class="finding-section">
          <div class="finding-label">What's happening</div>
          <div class="finding-text">${escapeHtml(f.explanation)}</div>
        </div>
        <div class="finding-section finding-tip">
          <div class="finding-label">💡 How to spot this yourself</div>
          <div class="finding-text">${escapeHtml(f.howToSpotIt)}</div>
        </div>
      </div>
    </div>
  `).join('');

  const linksHTML = links.length > 0
    ? '<div class="section"><div class="section-title">🔗 Links (' + links.length + ')</div>' +
      links.map(l => '<div class="link-item' + (l.mismatch ? ' link-mismatch' : '') + '"><span class="link-display">' + escapeHtml(l.display) + '</span><span class="link-domain">→ ' + escapeHtml(l.href) + (l.mismatch ? ' ⚠️ MISMATCH' : '') + '</span></div>').join('') + '</div>'
    : '';

  const attachWarn = highRisk.length > 0
    ? '<div class="attach-high-risk">⚠️ HIGH RISK ATTACHMENT: ' + escapeHtml(highRisk.join(', ')) + '<br>Do NOT open. Report to IT security immediately.</div>'
    : suspicious.length > 0
    ? '<div class="attach-suspicious">⚠️ SUSPICIOUS ATTACHMENT: ' + escapeHtml(suspicious.join(', ')) + '<br>Verify with sender before opening.</div>'
    : '';

  const combined = ((subject || '') + ' ' + (result.summary || '')).toLowerCase();
  const showWarn = ['sign in','verification code','one-time','otp','log in','verify your','reset your password','confirm your','your account','click here to'].some(kw => combined.includes(kw))
    || result.verdict === 'PHISHING' || result.phishing_score >= 60;

  document.getElementById('result-body').innerHTML =
    attachWarn +
    '<div class="verdict-card ' + vc + '"><span class="verdict-icon">' + vi + '</span><span class="verdict-label">' + escapeHtml(result.verdict) + '</span></div>' +
    '<div class="scores"><div class="score-item"><span class="score-label">Phishing Risk</span><span class="score-val">' + escapeHtml(String(result.phishing_score)) + '/100</span><div class="score-bar"><div class="score-fill phishing-fill" style="width:' + result.phishing_score + '%"></div></div></div><div class="score-item"><span class="score-label">Spam Score</span><span class="score-val">' + escapeHtml(String(result.spam_score)) + '/100</span><div class="score-bar"><div class="score-fill spam-fill" style="width:' + result.spam_score + '%"></div></div></div></div>' +
    '<div class="section"><div class="section-title">Summary</div><p>' + escapeHtml(result.summary) + '</p></div>' +
    (showWarn ? '<div class="warning-banner">⚠️ If you did not request this, do not click any links and <strong>report this to your IT security team immediately.</strong></div>' : '') +
    (findingsHTML ? '<div class="section"><div class="section-title">🔍 What We Found — tap each to learn more</div>' + findingsHTML + '</div>' : '') +
    linksHTML +
    (result.lesson ? '<div class="lesson"><div class="lesson-title">📚 Remember for next time</div><div class="lesson-text">' + escapeHtml(result.lesson) + '</div></div>' : '') +
    '<div class="section"><div class="section-title">✅ Suggested Action</div><p>' + escapeHtml(result.suggested_action) + '</p></div>' +
    '<div class="feedback-section" id="feedback-section"><div class="feedback-title">Was this analysis accurate?</div><div class="feedback-buttons"><button class="feedback-btn fb-false-positive" id="fb-fp">👎 False Positive</button><button class="feedback-btn fb-missed-threat" id="fb-mt">🚨 Missed Threat</button></div></div>';

  document.getElementById('fb-fp').addEventListener('click', () => showFeedbackForm('false_positive', result));
  document.getElementById('fb-mt').addEventListener('click', () => showFeedbackForm('missed_threat', result));
  resetBtn();
}

function showFeedbackForm(feedbackType, result) {
  const section = document.getElementById('feedback-section');
  const label = feedbackType === 'false_positive' ? 'This email was flagged but is actually safe' : 'This email is spam or phishing but was not caught';
  section.innerHTML = '<div class="feedback-title">' + label + '</div><textarea id="fb-comment" class="feedback-comment" placeholder="Optional: tell us more..." maxlength="500" rows="3"></textarea><div class="feedback-actions"><button class="feedback-btn fb-submit" id="fb-submit">Send Report</button><button class="feedback-btn fb-cancel" id="fb-cancel">Cancel</button></div>';
  document.getElementById('fb-submit').addEventListener('click', () => submitFeedback(feedbackType, result, (document.getElementById('fb-comment').value || '').trim()));
  document.getElementById('fb-cancel').addEventListener('click', resetFeedbackSection);
}

async function submitFeedback(feedbackType, result, comment) {
  const section = document.getElementById('feedback-section');
  section.innerHTML = '<div class="feedback-title" style="text-align:center;"><div class="spinner" style="margin:0 auto 6px;"></div>Sending report...</div>';
  const proxyUrl = storageGet('proxyUrl');
  const extToken = storageGet('extToken');
  if (!proxyUrl || !extToken) { section.innerHTML = '<div class="feedback-title" style="color:#a80000;">Extension not configured.</div>'; return; }
  const feedbackUrl = proxyUrl.replace(/\/analyze-email\/?$/, '/report-feedback');
  try {
    const item = Office.context.mailbox.item;
    const response = await fetch(feedbackUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token: extToken, feedbackType, originalVerdict: result.verdict, originalPhishingScore: result.phishing_score, originalSpamScore: result.spam_score, emailSubject: (item.subject || '').slice(0, 200), emailSender: item.from ? item.from.emailAddress.slice(0, 200) : '', userComment: comment })
    });
    section.innerHTML = response.ok
      ? '<div class="feedback-title" style="color:#107c10;">✅ Thank you! Report submitted.</div>'
      : '<div class="feedback-title" style="color:#a80000;">Failed to send. Please try again.</div>';
  } catch(e) { section.innerHTML = '<div class="feedback-title" style="color:#a80000;">Failed to send: ' + escapeHtml(e.message) + '</div>'; }
}

function resetFeedbackSection() {
  const section = document.getElementById('feedback-section');
  if (!section) return;
  const lastResult = window._oe_lastResult || {};
  section.innerHTML = '<div class="feedback-title">Was this analysis accurate?</div><div class="feedback-buttons"><button class="feedback-btn fb-false-positive" id="fb-fp">👎 False Positive</button><button class="feedback-btn fb-missed-threat" id="fb-mt">🚨 Missed Threat</button></div>';
  document.getElementById('fb-fp').addEventListener('click', () => showFeedbackForm('false_positive', lastResult));
  document.getElementById('fb-mt').addEventListener('click', () => showFeedbackForm('missed_threat', lastResult));
}
