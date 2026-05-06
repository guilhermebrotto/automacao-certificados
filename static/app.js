'use strict';

let sessionId        = null;
let emailSessionId   = null;
let pollTimer        = null;
let emailPollTimer   = null;
let lastMsgIdx       = 0;
let lastEmailMsgIdx  = 0;
let eligibleCount    = 0;

// ── File pickers ──────────────────────────────────────────────────────────────

function wireFilePicker(inputId, dropId, nameId, errorId) {
  const input = document.getElementById(inputId);
  const drop  = document.getElementById(dropId);

  input.addEventListener('change', () => {
    const file = input.files[0];
    if (!file) return;
    document.getElementById(nameId).textContent = file.name;
    drop.classList.add('is-filled');
    drop.classList.remove('is-error');
    clearFieldError(errorId, dropId);
  });

  // Drag-and-drop
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('is-filled'); });
  drop.addEventListener('dragleave', () => { if (!input.files[0]) drop.classList.remove('is-filled'); });
  drop.addEventListener('drop', e => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (!file) return;
    const dt = new DataTransfer();
    dt.items.add(file);
    input.files = dt.files;
    document.getElementById(nameId).textContent = file.name;
    drop.classList.add('is-filled');
    drop.classList.remove('is-error');
    clearFieldError(errorId, dropId);
  });
}

wireFilePicker('csv-input',  'csv-drop',  'csv-name',  'csv-error');
wireFilePicker('pptx-input', 'pptx-drop', 'pptx-name', 'pptx-error');

document.getElementById('event-name').addEventListener('input', () => {
  clearFieldError('event-error');
  document.getElementById('event-name').classList.remove('is-error');
});
document.getElementById('instructor').addEventListener('input', () => {
  clearFieldError('instructor-error');
  document.getElementById('instructor').classList.remove('is-error');
});

// ── Generate ──────────────────────────────────────────────────────────────────

async function handleGenerate() {
  const csvFile   = document.getElementById('csv-input').files[0];
  const pptxFile  = document.getElementById('pptx-input').files[0];
  const eventName = document.getElementById('event-name').value.trim();
  const instructor = document.getElementById('instructor').value.trim();

  let ok = true;
  if (!csvFile)    { setFieldError('csv-error',        'csv-drop',  'Selecione o arquivo CSV de participantes.'); ok = false; }
  if (!pptxFile)   { setFieldError('pptx-error',       'pptx-drop', 'Selecione o modelo PPTX.');                 ok = false; }
  if (!eventName)  { setTextError('event-error',    'event-name',   'Preencha o nome do evento.');                ok = false; }
  if (!instructor) { setTextError('instructor-error','instructor',   'Preencha o nome do ministrante.');           ok = false; }
  if (!ok) return;

  const btn = document.getElementById('btn-generate');
  btn.disabled = true;
  btn.innerHTML = spinner() + ' Iniciando…';

  const fd = new FormData();
  fd.append('csv',             csvFile);
  fd.append('pptx',            pptxFile);
  fd.append('event_name',      eventName);
  fd.append('instructor_name', instructor);

  try {
    const res  = await fetch('/gerar', { method: 'POST', body: fd });
    const data = await res.json();

    if (!res.ok) {
      handleServerErrors(data);
      btn.disabled = false;
      btn.innerHTML = generateBtnHTML();
      return;
    }

    sessionId  = data.session_id;
    lastMsgIdx = 0;

    document.getElementById('form-section').style.display = 'none';
    showSection('progress-section');
    document.getElementById('msg-feed').innerHTML = '';
    document.getElementById('progress-fill').style.width = '0%';
    document.getElementById('progress-pct').textContent  = '0%';
    document.getElementById('progress-label').textContent = 'Iniciando…';

    pollTimer = setInterval(pollStatus, 1000);

  } catch (err) {
    btn.disabled = false;
    btn.innerHTML = generateBtnHTML();
    showToast('Erro de conexão com o servidor.', 'err');
  }
}

async function pollStatus() {
  try {
    const res  = await fetch('/status/' + sessionId);
    if (!res.ok) return;
    const data = await res.json();

    // Update bar
    document.getElementById('progress-fill').style.width = data.progress + '%';
    document.getElementById('progress-pct').textContent  = Math.round(data.progress) + '%';
    if (data.total > 0)
      document.getElementById('progress-label').textContent = data.current + ' de ' + data.total + ' certificados';

    // Append new messages
    appendNewMessages('msg-feed', data.messages, 'lastMsgIdx');

    if (data.status === 'done') {
      clearInterval(pollTimer);
      eligibleCount = data.eligible_emails || 0;
      showResultSection(data);
    } else if (data.status === 'error') {
      clearInterval(pollTimer);
      showGenerationError(data.error);
    }
  } catch (_) { /* transient network error — keep polling */ }
}

function showResultSection(data) {
  const r = data.report;
  const allOk = r.failed === 0;
  const summaryEl = document.getElementById('result-summary');

  if (allOk) {
    summaryEl.innerHTML =
      '<div class="result-badge success">✅ ' + r.total + pluralize(r.total, ' certificado gerado', ' certificados gerados') + ' com sucesso!</div>' +
      '<div class="result-detail">Taxa de sucesso: ' + r.success_rate.toFixed(1) + '%</div>';
  } else {
    summaryEl.innerHTML =
      '<div class="result-badge warning">⚠️ ' + r.success + ' de ' + r.total + ' certificados gerados</div>' +
      '<div class="result-detail">Falhas: ' + esc(r.failed_names.join(', ')) + '</div>';
  }

  if (eligibleCount === 0) {
    const emailBtn = document.getElementById('btn-email');
    emailBtn.disabled = true;
    emailBtn.title = 'Nenhum participante com email @mindworks.com.br';
  }

  hideSection('progress-section');
  showSection('result-section');
}

function showGenerationError(msg) {
  document.getElementById('progress-title').textContent = 'Erro na Geração';
  const wrap = document.getElementById('progress-section');
  const existing = wrap.querySelector('.alert');
  if (existing) existing.remove();
  const el = document.createElement('div');
  el.className = 'alert alert-error';
  el.innerHTML =
    '<span>⚠</span>' +
    '<div class="alert-body">' + esc(msg) + '</div>' +
    '<span class="alert-close" onclick="handleReset()">↺ Tentar novamente</span>';
  wrap.appendChild(el);
}

// ── Download ──────────────────────────────────────────────────────────────────

function handleDownload() {
  if (sessionId) window.location = '/download/' + sessionId;
}

// ── Email ─────────────────────────────────────────────────────────────────────

async function handleEmail() {
  if (!sessionId) return;
  if (eligibleCount === 0) {
    showToast('Nenhum participante com email @mindworks.com.br.', 'warn');
    return;
  }

  const confirmed = confirm(
    'Enviar certificados por e-mail para ' + eligibleCount +
    pluralize(eligibleCount, ' participante', ' participantes') +
    ' com endereço @mindworks.com.br?\n\nParticipantes com outros domínios serão ignorados.'
  );
  if (!confirmed) return;

  const btn = document.getElementById('btn-email');
  btn.disabled = true;
  btn.innerHTML = spinner() + ' Enviando…';

  try {
    const res  = await fetch('/enviar/' + sessionId, { method: 'POST' });
    const data = await res.json();

    if (!res.ok) {
      showToast(data.error || 'Erro ao iniciar envio.', 'err');
      btn.disabled = false;
      btn.innerHTML = emailBtnHTML();
      return;
    }

    emailSessionId  = data.email_session_id;
    lastEmailMsgIdx = 0;

    const emailSection = document.getElementById('email-section');
    emailSection.style.display = 'block';
    document.getElementById('email-feed').innerHTML   = '';
    document.getElementById('email-result').innerHTML = '';

    emailPollTimer = setInterval(pollEmailStatus, 1000);

  } catch (err) {
    showToast('Erro de conexão.', 'err');
    btn.disabled = false;
    btn.innerHTML = emailBtnHTML();
  }
}

async function pollEmailStatus() {
  try {
    const res  = await fetch('/email-status/' + emailSessionId);
    if (!res.ok) return;
    const data = await res.json();

    appendNewMessages('email-feed', data.messages, 'lastEmailMsgIdx');

    if (data.status === 'done') {
      clearInterval(emailPollTimer);
      showEmailResult(data.report);
      const btn = document.getElementById('btn-email');
      btn.innerHTML = '✓ E-mails Enviados';
    } else if (data.status === 'error') {
      clearInterval(emailPollTimer);
      document.getElementById('email-result').innerHTML =
        '<div class="result-chip error">⚠ ' + esc(data.error) + '</div>';
      const btn = document.getElementById('btn-email');
      btn.disabled = false;
      btn.innerHTML = emailBtnHTML();
    }
  } catch (_) { /* ignore */ }
}

function showEmailResult(r) {
  const cls   = r.failed > 0 ? 'warning' : 'success';
  const icon  = r.failed > 0 ? '⚠️' : '✅';
  let   stats = icon + ' ' + r.sent + pluralize(r.sent, ' enviado', ' enviados');
  if (r.failed  > 0) stats += ' · ⚠ ' + r.failed + pluralize(r.failed, ' falha', ' falhas');
  if (r.skipped > 0) stats += ' · ' + r.skipped + ' ignorado' + (r.skipped !== 1 ? 's' : '');

  let note = '';
  if (r.skipped > 0)
    note = '<div class="result-chip-note">' + r.skipped +
           pluralize(r.skipped, ' participante ignorado', ' participantes ignorados') +
           ' (email fora do domínio @mindworks).</div>';

  document.getElementById('email-result').innerHTML =
    '<div class="result-chip ' + cls + '">' + stats + '</div>' + note;

  if (r.skipped > 0)
    showToast(r.skipped + pluralize(r.skipped, ' email ignorado', ' emails ignorados') + ' (domínio inválido).', 'warn');
}

// ── Reset ─────────────────────────────────────────────────────────────────────

function handleReset() {
  clearInterval(pollTimer);
  clearInterval(emailPollTimer);
  sessionId = emailSessionId = null;
  lastMsgIdx = lastEmailMsgIdx = 0;
  eligibleCount = 0;

  // Reset inputs
  document.getElementById('csv-input').value  = '';
  document.getElementById('pptx-input').value = '';
  document.getElementById('csv-name').textContent  = 'Clique para selecionar';
  document.getElementById('pptx-name').textContent = 'Clique para selecionar';
  document.getElementById('csv-drop').className  = 'file-drop';
  document.getElementById('pptx-drop').className = 'file-drop';
  document.getElementById('event-name').value = '';
  document.getElementById('instructor').value  = '';
  ['csv-error','pptx-error','event-error','instructor-error'].forEach(id => {
    document.getElementById(id).textContent = '';
  });

  // Reset progress card
  document.getElementById('progress-title').textContent = 'Gerando Certificados…';
  document.getElementById('progress-fill').style.width  = '0%';
  document.getElementById('progress-pct').textContent   = '0%';
  document.getElementById('progress-label').textContent = 'Iniciando…';
  document.getElementById('msg-feed').innerHTML = '';
  const old = document.getElementById('progress-section').querySelector('.alert');
  if (old) old.remove();

  // Reset result card
  document.getElementById('result-summary').innerHTML  = '';
  document.getElementById('email-section').style.display = 'none';
  document.getElementById('email-result').innerHTML    = '';
  document.getElementById('email-feed').innerHTML      = '';

  const generateBtn = document.getElementById('btn-generate');
  generateBtn.disabled = false;
  generateBtn.innerHTML = generateBtnHTML();

  const emailBtn = document.getElementById('btn-email');
  emailBtn.disabled = false;
  emailBtn.innerHTML = emailBtnHTML();

  hideSection('progress-section');
  hideSection('result-section');
  showSection('form-section');
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function appendNewMessages(feedId, messages, counterKey) {
  const feed    = document.getElementById(feedId);
  const icons   = { info: '✓', warning: '⚠', error: '✕' };
  const start   = counterKey === 'lastMsgIdx' ? lastMsgIdx : lastEmailMsgIdx;
  const newMsgs = messages.slice(start);
  if (counterKey === 'lastMsgIdx')      lastMsgIdx      = messages.length;
  else                                  lastEmailMsgIdx = messages.length;

  newMsgs.forEach(m => {
    const el = document.createElement('div');
    el.className = 'msg msg-' + m.level;
    el.innerHTML = '<span class="msg-icon">' + (icons[m.level] || '•') + '</span><span>' + esc(m.text) + '</span>';
    feed.appendChild(el);
  });
  if (newMsgs.length) feed.scrollTop = feed.scrollHeight;
}

function setFieldError(errId, dropId, msg) {
  document.getElementById(errId).textContent = msg;
  document.getElementById(dropId).classList.add('is-error');
}

function setTextError(errId, inputId, msg) {
  document.getElementById(errId).textContent = msg;
  document.getElementById(inputId).classList.add('is-error');
}

function clearFieldError(errId, dropId) {
  document.getElementById(errId).textContent = '';
  if (dropId) document.getElementById(dropId).classList.remove('is-error');
}

function handleServerErrors(data) {
  if (!data.errors) { showToast(data.error || 'Erro desconhecido.', 'err'); return; }
  const e = data.errors;
  if (e.csv)             setFieldError('csv-error',         'csv-drop',  e.csv);
  if (e.pptx)            setFieldError('pptx-error',        'pptx-drop', e.pptx);
  if (e.event_name)      setTextError('event-error',     'event-name',   e.event_name);
  if (e.instructor_name) setTextError('instructor-error', 'instructor',  e.instructor_name);
}

function showSection(id) { document.getElementById(id).style.display = 'block'; }
function hideSection(id) { document.getElementById(id).style.display = 'none';  }

function showToast(msg, type) {
  const c   = document.getElementById('toast-container');
  const el  = document.createElement('div');
  el.className = 'toast' + (type ? ' ' + type : '');
  el.textContent = msg;
  c.appendChild(el);
  setTimeout(() => {
    el.style.animation = 'toastOut .3s ease forwards';
    setTimeout(() => el.remove(), 300);
  }, 4200);
}

function esc(str) {
  const d = document.createElement('div');
  d.textContent = str;
  return d.innerHTML;
}

function pluralize(n, singular, plural) { return n === 1 ? singular : plural; }

function spinner() {
  return '<svg class="btn-svg" style="animation:spin 1s linear infinite" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg>';
}

function generateBtnHTML() {
  return 'Gerar Certificados';
}

function emailBtnHTML() {
  return '<svg class="btn-svg" viewBox="0 0 20 20" fill="currentColor"><path d="M3 4a2 2 0 0 0-2 2v1.161l8.441 4.221a1.25 1.25 0 0 0 1.118 0L19 7.162V6a2 2 0 0 0-2-2H3Z"/><path d="m19 8.839-7.77 3.885a2.75 2.75 0 0 1-2.46 0L1 8.839V14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V8.839Z"/></svg> Enviar por E-mail';
}

// CSS spin keyframe (injected once)
const spinStyle = document.createElement('style');
spinStyle.textContent = '@keyframes spin { to { transform: rotate(360deg); } }';
document.head.appendChild(spinStyle);
