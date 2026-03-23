/* Outlook Email Evaluator - Desktop Add-in (taskpane.js) */

Office.onReady(() => { initUI(); loadEmail(); });

function storageGet(key) { try { return localStorage.getItem(key) || ''; } catch(e) { return ''; } }
function storageSet(key, value) { try { localStorage.setItem(key, value); } catch(e) {} }

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
  populateSettings();
}

function populateSettings() {
  document.getElementById('api-key-input').value       = storageGet('apiKey');
  document.getElementById('tenant-domain-input').value = storageGet('tenantDomain');
  document.getElementById('custom-prompt-input').value = storageGet('customPrompt');
}

function saveSettings() {
  storageSet('apiKey',       document.getElementById('api-key-input').value.trim());
  storageSet('tenantDomain', document.getElementById('tenant-domain-input').value.trim());
  storageSet('customPrompt', document.getElementById('custom-prompt-input').value.trim());
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

async function analyzeEmail() {
  const apiKey = storageGet('apiKey');
  if (!apiKey) { showError('No API key set. Click the gear icon to add your Anthropic API key.'); return; }
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
  const tenantDomain     = storageGet('tenantDomain') || '';
  const senderEmail      = item.from ? item.from.emailAddress : '';
  const isExternal       = tenantDomain && senderEmail && !senderEmail.toLowerCase().endsWith('@' + tenantDomain.toLowerCase());
  const customPrompt     = storageGet('customPrompt') || '';
  const customPromptLine = customPrompt ? ('- Additional instructions: ' + customPrompt) : '';

  const prompt = buildPrompt({ subject, sender, body: bodyText.slice(0, 3000), links, attachNames, highRisk, suspicious, isExternal, tenantDomain, customPromptLine, utcString: new Date().toUTCString() });

  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
      body: JSON.stringify({ model: 'claude-sonnet-4-20250514', max_tokens: 1000, messages: [{ role: 'user', content: prompt }] })
    });
    if (!response.ok) { const err = await response.json(); showError('API ' + response.status + ': ' + (err.error?.message || JSON.stringify(err))); return; }
    const data = await response.json();
    const result = JSON.parse(data.content[0].text.trim().replace(/```json|```/g, '').trim());
    showResult(result, { subject, links, highRisk, suspicious });
  } catch (err) { showError('Request failed: ' + err.message); }
}

function extractLinks(html) {
  if (!html) return [];
  const doc = (new DOMParser()).parseFromString(html, 'text/html');
  const seen = new Set(); const links = [];
  doc.querySelectorAll('a[href]').forEach(a => {
    try {
      const displayText = (a.textContent || '').trim();
      let href = a.getAttribute('href') || '';
      if (href.includes('safelinks.protection.outlook.com') || href.includes('urldefense') || href.includes('trendmicro')) {
        try { const u = new URL(href); const p = u.searchParams.get('url') || u.searchParams.get('u'); if (p) href = decodeURIComponent(p); } catch(e) {}
      }
      if (!href || href.startsWith('mailto:') || href.startsWith('#') || href.length < 10) return;
      let hrefDomain = ''; try { hrefDomain = new URL(href).hostname.toLowerCase(); } catch(e) { return; }
      if (seen.has(hrefDomain)) return; seen.add(hrefDomain);
      let displayDomain = '';
      const urlPattern = displayText.match(/(?:https?:\/\/|www\.)([\w.-]+)/i);
      if (urlPattern) { try { displayDomain = new URL(displayText.startsWith('http') ? displayText : 'https://' + displayText).hostname.toLowerCase(); } catch(e) { displayDomain = urlPattern[1].toLowerCase(); } }
      const mismatch = displayDomain && hrefDomain && !hrefDomain.includes(displayDomain.replace(/^www\./, '')) && !displayDomain.includes(hrefDomain.replace(/^www\./, ''));
      links.push({ display: displayText.slice(0, 80) || '(no text)', href: hrefDomain, mismatch });
    } catch(e) {}
  });
  return links.slice(0, 20);
}

function buildPrompt({ subject, sender, body, links, attachNames, highRisk, suspicious, isExternal, tenantDomain, customPromptLine, utcString }) {
  const attachInfo     = attachNames.length > 0 ? attachNames.join(', ') : '(none)';
  const highRiskNote   = highRisk.length > 0 ? 'CRITICAL: HIGH RISK attachment(s) detected: ' + highRisk.join(', ') + '. You MUST set verdict to PHISHING, phishing_score >= 90, and suggested_action MUST include: Do NOT open this attachment. Report to IT security immediately.' : '';
  const suspiciousNote = (suspicious.length > 0 && highRisk.length === 0) ? 'WARNING: SUSPICIOUS attachment(s): ' + suspicious.join(', ') + '. Set phishing_score >= 60. Do not open unless certain of origin.' : '';
  const linksText      = links.length > 0 ? links.map(l => ' - Display: "' + l.display + '" -> Real domain: ' + l.href + (l.mismatch ? ' WARNING: DOMAIN MISMATCH' : '')).join('\n') : ' (No links found)';
  return 'You are a cybersecurity expert. Analyze this email and respond ONLY with a JSON object.\n\nCONTEXT:\n- Date: ' + utcString + '\n- Org domain: ' + (tenantDomain||'unknown') + '\n- Sender: ' + sender + '\n- External: ' + (isExternal?'YES':'NO') + '\n- SharePoint/OneDrive from ' + (tenantDomain||'unknown') + ' = INTERNAL, never flag\n' + customPromptLine + '\n\nRULES:\n1. Never free-pass based on sender domain alone.\n2. Flag suspicious content/urgency/credential requests regardless of sender.\n3. Payments/access changes: suggested_action MUST include "Verify through official channels other than email."\n4. Login link/OTP/security alert: suggested_action MUST include "If you did not request this, do not click links and report to IT security immediately."\n5. Verification code: suggested_action MUST include "Never share this code with anyone."\n\nSubject: ' + subject + '\nFrom: ' + sender + '\nBody:\n' + body + '\n\nAttachments: ' + attachInfo + '\n' + highRiskNote + '\n' + suspiciousNote + '\n\nLINKS (decoded):\n' + linksText + '\n\nRespond with ONLY this JSON:\n{"verdict":"SAFE"|"SUSPICIOUS"|"SPAM"|"PHISHING","phishing_score":<0-100>,"spam_score":<0-100>,"reasons":[<strings>],"suggested_action":"<string>","summary":"<1-2 sentences>"}';
}

function setLoading() {
  document.getElementById('result-body').innerHTML = '<div class="loading"><div class="spinner"></div><span>Analyzing email...</span></div>';
  document.getElementById('analyze-btn').disabled = true;
  document.getElementById('analyze-btn').textContent = 'Analyzing...';
}

function showError(msg) {
  document.getElementById('result-body').innerHTML = '<div class="error">Warning: ' + msg + '</div>';
  resetBtn();
}

function resetBtn() {
  const btn = document.getElementById('analyze-btn');
  btn.disabled = false;
  btn.textContent = 'Analyze Email';
}

function showResult(result, { subject, links, highRisk, suspicious }) {
  const vc = { SAFE:'verdict-safe', SUSPICIOUS:'verdict-suspicious', SPAM:'verdict-spam', PHISHING:'verdict-phishing' }[result.verdict] || 'verdict-suspicious';
  const vi = { SAFE:'OK', SUSPICIOUS:'!!', SPAM:'SPAM', PHISHING:'PHISHING' }[result.verdict] || '!!';
  const reasonsHTML = (result.reasons||[]).map(r => '<li>'+r+'</li>').join('');
  const linksHTML   = links.length > 0 ? '<div class="section"><div class="section-title">Links ('+links.length+')</div>'+links.map(l=>'<div class="link-item'+(l.mismatch?' link-mismatch':'')+'"><span class="link-display">'+l.display+'</span><span class="link-domain">-> '+l.href+(l.mismatch?' MISMATCH':'')+'</span></div>').join('')+'</div>' : '';
  const attachWarn  = highRisk.length>0 ? '<div class="attach-high-risk">HIGH RISK ATTACHMENT: '+highRisk.join(', ')+'<br>Do NOT open. Report to IT security immediately.</div>' : suspicious.length>0 ? '<div class="attach-suspicious">SUSPICIOUS ATTACHMENT: '+suspicious.join(', ')+'<br>Verify with sender before opening.</div>' : '';
  const combined    = (subject+' '+(result.summary||'')).toLowerCase();
  const showWarn    = ['sign in','verification code','one-time','otp','log in','verify your','reset your password','your account'].some(kw=>combined.includes(kw)) || result.verdict==='PHISHING' || result.phishing_score>=60;

  document.getElementById('result-body').innerHTML =
    attachWarn +
    '<div class="verdict-card '+vc+'"><span class="verdict-icon">'+vi+'</span><span class="verdict-label">'+result.verdict+'</span></div>' +
    '<div class="scores"><div class="score-item"><span class="score-label">Phishing Risk</span><span class="score-val">'+result.phishing_score+'/100</span><div class="score-bar"><div class="score-fill phishing-fill" style="width:'+result.phishing_score+'%"></div></div></div><div class="score-item"><span class="score-label">Spam Score</span><span class="score-val">'+result.spam_score+'/100</span><div class="score-bar"><div class="score-fill spam-fill" style="width:'+result.spam_score+'%"></div></div></div></div>' +
    '<div class="section"><div class="section-title">Summary</div><p>'+result.summary+'</p></div>' +
    (reasonsHTML ? '<div class="section"><div class="section-title">Why it\'s suspicious</div><ul>'+reasonsHTML+'</ul></div>' : '') +
    linksHTML +
    (showWarn ? '<div class="warning-banner">If you did not request this, do not click any links and report to your IT security team immediately.</div>' : '') +
    '<div class="section"><div class="section-title">Suggested Action</div><p>'+result.suggested_action+'</p></div>';

  resetBtn();
}
