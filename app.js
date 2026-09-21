/* 寮 緊急時対応ガイド  app.js */
'use strict';

/* ---------- 端末内の保存 ---------- */
const LS = {
  get(k, d) { try { const v = localStorage.getItem('emg.' + k); return v === null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem('emg.' + k, JSON.stringify(v)); return true; } catch (e) { return false; } },
  del(k) { try { localStorage.removeItem('emg.' + k); } catch (e) { /* noop */ } }
};

/* ---------- 画面の文言 ---------- */
const UI = {
  ja: {
    setupTitle: 'はじめに', setupLead: '研修で伝えられたパスコードと、あなたの名前を入力してください。入力は最初の1回だけです。',
    rsName: 'あなたの名前', passcode: 'パスコード', start: '使いはじめる', loading: '読み込んでいます…',
    wrongPass: 'パスコードが違います。研修で伝えられたパスコードを確認してください。',
    needName: '名前とパスコードを入力してください。',
    noData: 'データがまだ公開されていません。担当職員に連絡してください。',
    offlineNoData: '初回は通信できる場所で開いてください。一度読み込めば、以降は通信がなくても使えます。',
    home: 'ホーム', contacts: '電話帳', settings: '設定',
    offlineReady: 'オフラインでも使えます', offlineNotReady: 'オフライン準備が未完了です（通信できる場所で一度開いてください）',
    online: '通信あり', offline: '通信なし',
    dataVer: 'データの版', updated: '更新', lastCheck: '最終確認', never: '未確認',
    chooseScene: '起きていることを選んでください',
    installTitle: 'ホーム画面に追加してください',
    installIOS: 'Safariの共有ボタン（□と↑）から「ホーム画面に追加」を選ぶと、通信がなくても開けるようになります。',
    installOther: 'ブラウザのメニューから「ホーム画面に追加」または「アプリをインストール」を選んでください。',
    back: '戻る', restart: '最初から', restartConfirm: '最初からやり直しますか？（この対応は記録されません）',
    doNot: 'やらないこと', memo: 'メモ（任意）', memoPh: '部屋番号、時刻、状況など',
    finish: '記録して終了', logged: '記録しました', loggedPending: '記録しました（通信が戻ったら送信します）',
    pending: '未送信の記録', sendNow: '今すぐ送信', sent: '送信しました', sendFail: '送信できませんでした。通信が戻ったら自動で送信します。',
    checkUpdate: '更新を確認', upToDate: '最新の状態です', dataUpdated: '最新のデータに更新しました', updateFail: '更新を確認できませんでした',
    passChanged: 'パスコードが変更されました。新しいパスコードを入力してください。', reenter: '入力する',
    newApp: 'アプリの新しい版があります。', reload: '更新',
    yourName: '名前', save: '保存', saved: '保存しました', language: '表示言語',
    reenterPass: 'パスコードを入力し直す', reset: 'この端末のデータを削除',
    resetConfirm: 'この端末に保存されたデータと未送信の記録を削除します。よろしいですか？',
    storage: 'データの保持', persistOk: '端末に保持されます', persistNg: '長期間開かないと端末から消える可能性があります。月に1回は開いてください。',
    call: '電話する', noContacts: '連絡先がありません',
    pathLabel: 'たどった経路', items: '件'
  },
  en: {
    setupTitle: 'Getting started', setupLead: 'Enter the passcode you were given at training and your name. You only need to do this once.',
    rsName: 'Your name', passcode: 'Passcode', start: 'Start', loading: 'Loading…',
    wrongPass: 'Wrong passcode. Check the passcode you were given at training.',
    needName: 'Enter your name and the passcode.',
    noData: 'No data has been published yet. Contact the staff in charge.',
    offlineNoData: 'Open the app with an internet connection the first time. After that it works offline.',
    home: 'Home', contacts: 'Contacts', settings: 'Settings',
    offlineReady: 'Works offline', offlineNotReady: 'Offline setup is not complete (open once with a connection)',
    online: 'Online', offline: 'Offline',
    dataVer: 'Data version', updated: 'Updated', lastCheck: 'Last checked', never: 'Never',
    chooseScene: 'What is happening?',
    installTitle: 'Add this app to your home screen',
    installIOS: 'In Safari, tap the Share button and choose "Add to Home Screen" so it opens without a connection.',
    installOther: 'From the browser menu, choose "Add to Home screen" or "Install app".',
    back: 'Back', restart: 'Start over', restartConfirm: 'Start over? This response will not be recorded.',
    doNot: 'Do not', memo: 'Note (optional)', memoPh: 'Room number, time, situation, etc.',
    finish: 'Save record and finish', logged: 'Record saved', loggedPending: 'Record saved (it will be sent when you are online)',
    pending: 'Unsent records', sendNow: 'Send now', sent: 'Sent', sendFail: 'Could not send. It will be sent automatically when online.',
    checkUpdate: 'Check for updates', upToDate: 'Up to date', dataUpdated: 'Updated to the latest data', updateFail: 'Could not check for updates',
    passChanged: 'The passcode has changed. Enter the new passcode.', reenter: 'Enter',
    newApp: 'A new version of the app is available.', reload: 'Update',
    yourName: 'Name', save: 'Save', saved: 'Saved', language: 'Language',
    reenterPass: 'Re-enter passcode', reset: 'Delete data on this device',
    resetConfirm: 'Delete the data and unsent records stored on this device?',
    storage: 'Data storage', persistOk: 'Kept on this device', persistNg: 'May be removed if unused for a long time. Open the app at least once a month.',
    call: 'Call', noContacts: 'No contacts',
    pathLabel: 'Path taken', items: ''
  }
};

/* ---------- 状態 ---------- */
let lang = LS.get('lang', (navigator.language || 'ja').startsWith('ja') ? 'ja' : 'en');
let flow = null;          // 復号したデータ
let needPass = false;     // 新しいパスコードが必要
let persisted = null;     // 永続化の結果
let tab = 'home';
let run = null;           // { scene, stack:[id], picks:[index], checks:{id:[bool]}, memo }
let waitingSW = null;

const $ = (s) => document.querySelector(s);
const t = (k) => (UI[lang] && UI[lang][k]) || UI.ja[k] || k;
const tx = (o) => (o && (o[lang] || o.ja)) || '';
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const lines = (s) => String(s || '').split('\n').map((x) => x.trim()).filter(Boolean);
const isStandalone = () => window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

function fmtDate(iso) {
  if (!iso) return t('never');
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  const p = (n) => String(n).padStart(2, '0');
  return d.getFullYear() + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
}

function uuid() {
  if (crypto.randomUUID) return crypto.randomUUID();
  const b = crypto.getRandomValues(new Uint8Array(16));
  return Array.from(b, (x) => x.toString(16).padStart(2, '0')).join('');
}

let toastTimer = null;
function toast(msg) {
  const el = $('#toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 3200);
}

/* ---------- 復号（GAS の encryptPayload_ と対になる処理） ---------- */
const b64d = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

function concatBytes(...arrs) {
  const out = new Uint8Array(arrs.reduce((n, a) => n + a.length, 0));
  let o = 0;
  arrs.forEach((a) => { out.set(a, o); o += a.length; });
  return out;
}

async function decryptFlow(enc, pass) {
  if (!enc || enc.v !== 1) throw new Error('FORMAT');
  const te = new TextEncoder();
  const salt = b64d(enc.salt), nonce = b64d(enc.nonce), ct = b64d(enc.ct), mac = b64d(enc.mac);
  const base = await crypto.subtle.importKey('raw', te.encode(pass), 'PBKDF2', false, ['deriveBits']);
  const bits = new Uint8Array(await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations: enc.iter }, base, 512));
  const hk = (raw) => crypto.subtle.importKey('raw', raw, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign', 'verify']);
  const encKey = await hk(bits.slice(0, 32));
  const macKey = await hk(bits.slice(32, 64));
  const ok = await crypto.subtle.verify('HMAC', macKey, mac, concatBytes(salt, nonce, ct));
  if (!ok) throw new Error('BADPASS');
  const blocks = Math.ceil(ct.length / 32);
  const ks = await Promise.all(Array.from({ length: blocks }, (_, i) => {
    const m = new Uint8Array(nonce.length + 4);
    m.set(nonce);
    new DataView(m.buffer).setUint32(nonce.length, i);
    return crypto.subtle.sign('HMAC', encKey, m).then((b) => new Uint8Array(b));
  }));
  const out = new Uint8Array(ct.length);
  for (let i = 0; i < ct.length; i++) out[i] = ct[i] ^ ks[i >> 5][i & 31];
  return JSON.parse(new TextDecoder().decode(out));
}

/* ---------- 通信 ---------- */
async function fetchLatest() {
  const vr = await fetch('data/version.json?t=' + Date.now(), { cache: 'no-store' });
  if (vr.status === 404) throw new Error('NODATA');
  if (!vr.ok) throw new Error('NET');
  const v = await vr.json();
  const er = await fetch('data/flow.enc.json?v=' + encodeURIComponent(v.version), { cache: 'no-store' });
  if (er.status === 404) throw new Error('NODATA');
  if (!er.ok) throw new Error('NET');
  return { v, enc: await er.json() };
}

async function checkUpdate(manual) {
  if (!navigator.onLine) { if (manual) toast(t('offline')); return; }
  try {
    const vr = await fetch('data/version.json?t=' + Date.now(), { cache: 'no-store' });
    if (!vr.ok) throw new Error('NET');
    const v = await vr.json();
    LS.set('lastCheck', new Date().toISOString());
    if (flow && v.version === LS.get('ver', null) && !needPass) {
      if (manual) toast(t('upToDate'));
      refreshChrome();
      return;
    }
    const er = await fetch('data/flow.enc.json?v=' + encodeURIComponent(v.version), { cache: 'no-store' });
    if (!er.ok) throw new Error('NET');
    const enc = await er.json();
    const data = await decryptFlow(enc, LS.get('pass', ''));
    LS.set('enc', enc); LS.set('ver', v.version); LS.set('updatedAt', v.updatedAt);
    flow = data; needPass = false;
    hideBanner('pass');
    toast(t('dataUpdated'));
    if (!run) render();
    else refreshChrome();
  } catch (e) {
    if (e.message === 'BADPASS') { needPass = true; showPassBanner(); }
    else if (manual) toast(t('updateFail'));
    refreshChrome();
  }
}

async function sendLogs(manual) {
  const logs = LS.get('logs', []);
  if (!logs.length) return;
  if (!flow || !flow.meta || !flow.meta.logUrl || !navigator.onLine) { if (manual) toast(t('sendFail')); return; }
  try {
    const res = await fetch(flow.meta.logUrl, { method: 'POST', body: JSON.stringify({ key: flow.meta.logKey, logs }) });
    const j = await res.json();
    if (!j.ok) throw new Error(j.error || 'NG');
    const sentIds = new Set(logs.map((l) => l.id));
    LS.set('logs', LS.get('logs', []).filter((l) => !sentIds.has(l.id)));
    if (manual) toast(t('sent'));
  } catch (e) {
    if (manual) toast(t('sendFail'));
  }
  if (tab === 'settings' && !run) render();
}

/* ---------- お知らせ ---------- */
function showBanner(kind, msg, btnLabel, onClick) {
  const b = $('#banner');
  b.dataset.kind = kind;
  b.innerHTML = '<span>' + esc(msg) + '</span><button type="button">' + esc(btnLabel) + '</button>';
  b.querySelector('button').onclick = onClick;
  b.hidden = false;
}
function hideBanner(kind) {
  const b = $('#banner');
  if (!kind || b.dataset.kind === kind) b.hidden = true;
}
function showPassBanner() {
  showBanner('pass', t('passChanged'), t('reenter'), () => { run = null; renderSetup(true); });
}

/* ---------- 共通の表示更新 ---------- */
function refreshChrome() {
  $('#appTitle').textContent = flow ? tx(flow.meta.title) : (lang === 'ja' ? '緊急時対応ガイド' : 'Emergency Guide');
  document.title = $('#appTitle').textContent;
  document.documentElement.lang = lang;
  $('#langBtn').textContent = lang === 'ja' ? 'EN' : '日本語';
  $('#tabHome').textContent = t('home');
  $('#tabContacts').textContent = t('contacts');
  $('#tabSettings').textContent = t('settings');
  document.querySelectorAll('.tabs button').forEach((b) => {
    if (b.dataset.tab === tab && !run) b.setAttribute('aria-current', 'page'); else b.removeAttribute('aria-current');
  });
  const st = $('#statusBox');
  if (st) st.innerHTML = statusHtml();
}

function statusHtml() {
  const ready = !!(flow && navigator.serviceWorker && navigator.serviceWorker.controller);
  return '<div class="row"><span class="dot ' + (ready ? 'ok' : 'ng') + '"></span>' + esc(ready ? t('offlineReady') : t('offlineNotReady')) + '</div>' +
    '<div class="row"><span class="dot ' + (navigator.onLine ? 'ok' : '') + '"></span>' + esc(navigator.onLine ? t('online') : t('offline')) +
    ' ／ ' + esc(t('updated')) + ' ' + esc(fmtDate(LS.get('updatedAt', ''))) + '</div>';
}

function render() {
  $('#tabs').hidden = !flow || !LS.get('rs', '');
  if (!flow || !LS.get('rs', '')) { renderSetup(false); return; }
  if (run) renderNode();
  else if (tab === 'contacts') renderContacts();
  else if (tab === 'settings') renderSettings();
  else renderHome();
  refreshChrome();
}

function setView(html) {
  const v = $('#view');
  v.innerHTML = html;
  window.scrollTo(0, 0);
  v.focus({ preventScroll: true });
}

/* ---------- 初回設定 ---------- */
function renderSetup(passOnly) {
  $('#tabs').hidden = true;
  refreshChrome();
  setView(
    '<h2 class="sec">' + esc(t('setupTitle')) + '</h2>' +
    '<p class="lead">' + esc(t('setupLead')) + '</p>' +
    '<p class="err" id="setupErr" hidden></p>' +
    '<div class="field"><label for="fName">' + esc(t('rsName')) + '</label><input id="fName" autocomplete="name" value="' + esc(LS.get('rs', '')) + '"' + (passOnly ? ' readonly' : '') + '></div>' +
    '<div class="field"><label for="fPass">' + esc(t('passcode')) + '</label><input id="fPass" type="password" autocomplete="current-password"></div>' +
    '<button class="btn-main" id="fGo" type="button">' + esc(t('start')) + '</button>' +
    installHint()
  );
  const go = $('#fGo'), err = $('#setupErr');
  const submit = async () => {
    const name = $('#fName').value.trim(), pass = $('#fPass').value;
    err.hidden = true;
    if (!name || !pass) { err.textContent = t('needName'); err.hidden = false; return; }
    go.disabled = true; go.textContent = t('loading');
    try {
      let enc = LS.get('enc', null), ver = LS.get('ver', null), upd = LS.get('updatedAt', null);
      if (navigator.onLine) {
        try {
          const r = await fetchLatest();
          enc = r.enc; ver = r.v.version; upd = r.v.updatedAt;
          LS.set('lastCheck', new Date().toISOString());
        } catch (e) { if (!enc) throw e; }
      }
      if (!enc) throw new Error(navigator.onLine ? 'NODATA' : 'OFFLINE');
      flow = await decryptFlow(enc, pass);
      LS.set('enc', enc); LS.set('ver', ver); LS.set('updatedAt', upd);
      LS.set('pass', pass); LS.set('rs', name);
      needPass = false; hideBanner('pass');
      tab = 'home'; run = null;
      render();
      sendLogs(false);
    } catch (e) {
      const m = e.message === 'BADPASS' ? t('wrongPass') : e.message === 'NODATA' ? t('noData') : t('offlineNoData');
      err.textContent = m; err.hidden = false;
      go.disabled = false; go.textContent = t('start');
    }
  };
  go.onclick = submit;
  $('#fPass').addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); });
}

function installHint() {
  if (isStandalone()) return '';
  return '<div class="hint"><b>' + esc(t('installTitle')) + '</b><br>' + esc(isIOS() ? t('installIOS') : t('installOther')) + '</div>';
}

/* ---------- ホーム ---------- */
function renderHome() {
  const other = lang === 'ja' ? 'en' : 'ja';
  const plates = flow.top.map((p, i) =>
    '<button type="button" class="plate ' + esc(p.color) + '" data-i="' + i + '">' +
    '<span class="mark" aria-hidden="true"></span><span><span class="l1">' + esc(tx(p.label)) + '</span>' +
    (p.label[other] ? '<span class="l2" lang="' + other + '">' + esc(p.label[other]) + '</span>' : '') + '</span></button>'
  ).join('');
  setView(
    '<div class="status" id="statusBox">' + statusHtml() + '</div>' +
    '<h2 class="sec">' + esc(t('chooseScene')) + '</h2>' +
    '<div class="plates">' + plates + '</div>' + installHint()
  );
  document.querySelectorAll('.plate').forEach((b) => {
    b.onclick = () => {
      const p = flow.top[Number(b.dataset.i)];
      run = { scene: p.label, stack: [p.start], picks: [], checks: {}, memo: '' };
      render();
    };
  });
}

/* ---------- フロー ---------- */
function renderNode() {
  const id = run.stack[run.stack.length - 1];
  const n = flow.nodes[id];
  if (!n) { run = null; render(); return; }

  const trail = run.stack.slice(0, -1).map((sid, i) => {
    const sn = flow.nodes[sid], pick = sn && sn.choices[run.picks[i]];
    return '<li>' + esc(sid) + (pick ? '：' + esc(tx(pick.label)) : '') + '</li>';
  }).join('');

  const dl = lines(tx(n.detail));
  let body = '';
  if (n.type === 'a' && dl.length) {
    const st = run.checks[id] || (run.checks[id] = dl.map(() => false));
    body += '<ul class="checks">' + dl.map((l, i) =>
      '<li><label><input type="checkbox" data-c="' + i + '"' + (st[i] ? ' checked' : '') + '><span>' + esc(l) + '</span></label></li>').join('') + '</ul>';
  } else if (dl.length) {
    body += '<ul class="lines">' + dl.map((l) => '<li>' + esc(l) + '</li>').join('') + '</ul>';
  }
  if (tx(n.dont)) body += '<div class="dont"><small>' + esc(t('doNot')) + '</small>' + esc(tx(n.dont)) + '</div>';
  if (n.contacts.length) body += '<div class="calls">' + n.contacts.map(callHtml).join('') + '</div>';

  let tail = '';
  if (n.choices.length) {
    tail += '<div class="choices">' + n.choices.map((c, i) =>
      '<button type="button" class="choice" data-k="' + i + '">' + esc(tx(c.label)) + '</button>').join('') + '</div>';
  }
  if (n.type === 'a') {
    tail += '<div class="finish"><label for="memo" style="font-weight:700">' + esc(t('memo')) + '</label>' +
      '<textarea id="memo" placeholder="' + esc(t('memoPh')) + '">' + esc(run.memo) + '</textarea>' +
      '<button type="button" class="btn-main" id="finish">' + esc(t('finish')) + '</button></div>';
  }

  setView(
    '<div class="flowbar"><button type="button" class="btn-ghost" id="back">' + esc(t('back')) + '</button>' +
    '<button type="button" class="btn-ghost" id="restart">' + esc(t('restart')) + '</button></div>' +
    '<p class="sub" style="margin:0 0 6px;color:var(--muted);font-size:14px">' + esc(tx(run.scene)) + '</p>' +
    (trail ? '<ol class="trail" aria-label="' + esc(t('pathLabel')) + '">' + trail + '</ol>' : '') +
    '<section class="node ' + esc(n.level) + '"><span class="idplate">' + esc(id) + '</span>' +
    '<h1>' + esc(tx(n.text)) + '</h1>' + body + '</section>' + tail
  );

  $('#back').onclick = () => {
    if (run.stack.length <= 1) { run = null; render(); return; }
    run.stack.pop(); run.picks.pop();
    render();
  };
  $('#restart').onclick = () => { if (confirm(t('restartConfirm'))) { run = null; tab = 'home'; render(); } };
  document.querySelectorAll('.checks input').forEach((cb) => {
    cb.onchange = () => { run.checks[id][Number(cb.dataset.c)] = cb.checked; };
  });
  document.querySelectorAll('.choice').forEach((b) => {
    b.onclick = () => {
      const k = Number(b.dataset.k);
      run.picks.push(k);
      run.stack.push(n.choices[k].next);
      render();
    };
  });
  const memo = $('#memo');
  if (memo) memo.oninput = () => { run.memo = memo.value; };
  const fin = $('#finish');
  if (fin) fin.onclick = () => finishRun(id);
}

function callHtml(cid) {
  const c = flow.contacts[cid];
  if (!c) return '';
  const href = 'tel:' + String(c.tel).replace(/[^\d#*+]/g, '').replace(/#/g, '%23');
  return '<a class="call" href="' + esc(href) + '"><span class="n">' + esc(tx(c.name)) +
    (tx(c.hours) ? '<small>' + esc(tx(c.hours)) + '</small>' : '') + '</span><span class="t">' + esc(c.tel) + '</span></a>';
}

function finishRun(endId) {
  const pathJa = run.stack.map((sid, i) => {
    const sn = flow.nodes[sid], pick = sn && sn.choices[run.picks[i]];
    return sid + (pick ? '「' + (pick.label.ja || pick.label.en) + '」' : '');
  }).join(' → ');
  const done = [], notDone = [];
  run.stack.forEach((sid) => {
    const sn = flow.nodes[sid];
    if (!sn || sn.type !== 'a') return;
    const ls = lines(sn.detail.ja || sn.detail.en);
    const st = run.checks[sid] || [];
    ls.forEach((l, i) => (st[i] ? done : notDone).push(sid + ': ' + l));
  });
  const logs = LS.get('logs', []);
  logs.push({
    id: uuid(), at: new Date().toISOString(), rs: LS.get('rs', ''),
    scenario: run.scene.ja || run.scene.en, path: pathJa, end: endId,
    done: done.join('\n'), notDone: notDone.join('\n'), memo: run.memo || '',
    version: LS.get('ver', '')
  });
  LS.set('logs', logs);
  run = null; tab = 'home';
  render();
  toast(navigator.onLine ? t('logged') : t('loggedPending'));
  sendLogs(false);
}

/* ---------- 電話帳 ---------- */
function renderContacts() {
  const ids = flow.contactOrder || Object.keys(flow.contacts);
  setView('<h2 class="sec">' + esc(t('contacts')) + '</h2>' +
    (ids.length ? '<div class="calls">' + ids.map(callHtml).join('') + '</div>' : '<p class="lead">' + esc(t('noContacts')) + '</p>'));
}

/* ---------- 設定 ---------- */
function renderSettings() {
  const pend = LS.get('logs', []).length;
  setView(
    '<h2 class="sec">' + esc(t('settings')) + '</h2>' +
    '<div class="panel"><div class="status" id="statusBox" style="margin:0">' + statusHtml() + '</div>' +
    '<dl class="kv" style="margin-top:10px"><dt>' + esc(t('dataVer')) + '</dt><dd>' + esc(LS.get('ver', '-')) + '</dd>' +
    '<dt>' + esc(t('lastCheck')) + '</dt><dd>' + esc(fmtDate(LS.get('lastCheck', ''))) + '</dd>' +
    '<dt>' + esc(t('storage')) + '</dt><dd>' + esc(persisted === true ? t('persistOk') : t('persistNg')) + '</dd></dl>' +
    '<div class="row-btns"><button type="button" class="btn-ghost" id="sUpd">' + esc(t('checkUpdate')) + '</button></div></div>' +

    '<div class="panel"><h3>' + esc(t('pending')) + '：' + pend + esc(t('items')) + '</h3>' +
    '<div class="row-btns"><button type="button" class="btn-ghost" id="sSend"' + (pend ? '' : ' disabled') + '>' + esc(t('sendNow')) + '</button></div></div>' +

    '<div class="panel"><div class="field" style="margin-bottom:8px"><label for="sName">' + esc(t('yourName')) + '</label>' +
    '<input id="sName" value="' + esc(LS.get('rs', '')) + '"></div>' +
    '<div class="row-btns"><button type="button" class="btn-ghost" id="sSave">' + esc(t('save')) + '</button>' +
    '<button type="button" class="btn-ghost" id="sPass">' + esc(t('reenterPass')) + '</button></div></div>' +

    '<div class="panel"><div class="row-btns" style="margin:0"><button type="button" class="btn-ghost danger" id="sReset">' + esc(t('reset')) + '</button></div></div>'
  );
  $('#sUpd').onclick = () => checkUpdate(true);
  $('#sSend').onclick = () => sendLogs(true);
  $('#sSave').onclick = () => { const v = $('#sName').value.trim(); if (v) { LS.set('rs', v); toast(t('saved')); } };
  $('#sPass').onclick = () => renderSetup(true);
  $('#sReset').onclick = () => {
    if (!confirm(t('resetConfirm'))) return;
    ['enc', 'ver', 'updatedAt', 'pass', 'rs', 'logs', 'lastCheck'].forEach(LS.del);
    flow = null; run = null; tab = 'home';
    render();
  };
}

/* ---------- Service Worker ---------- */
function registerSW() {
  if (!('serviceWorker' in navigator)) return;
  navigator.serviceWorker.register('sw.js').then((reg) => {
    const offer = (w) => {
      waitingSW = w;
      showBanner('app', t('newApp'), t('reload'), () => { if (waitingSW) waitingSW.postMessage('SKIP_WAITING'); });
    };
    if (reg.waiting && navigator.serviceWorker.controller) offer(reg.waiting);
    reg.addEventListener('updatefound', () => {
      const w = reg.installing;
      if (!w) return;
      w.addEventListener('statechange', () => {
        if (w.state === 'installed' && navigator.serviceWorker.controller) offer(w);
        if (w.state === 'activated') refreshChrome();
      });
    });
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') reg.update().catch(() => {}); });
  }).catch(() => {});
  let reloading = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (waitingSW && !reloading) { reloading = true; location.reload(); }
    else refreshChrome();
  });
}

/* ---------- 起動 ---------- */
async function boot() {
  $('#langBtn').onclick = () => { lang = lang === 'ja' ? 'en' : 'ja'; LS.set('lang', lang); render(); if (needPass) showPassBanner(); };
  document.querySelectorAll('.tabs button').forEach((b) => { b.onclick = () => { tab = b.dataset.tab; run = null; render(); }; });
  window.addEventListener('online', () => { refreshChrome(); checkUpdate(false); sendLogs(false); });
  window.addEventListener('offline', refreshChrome);
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && flow) { checkUpdate(false); sendLogs(false); }
  });

  registerSW();
  if (navigator.storage && navigator.storage.persist) {
    try { persisted = (await navigator.storage.persisted()) || (await navigator.storage.persist()); } catch (e) { persisted = false; }
  }

  const enc = LS.get('enc', null), pass = LS.get('pass', null);
  if (enc && pass) {
    try { flow = await decryptFlow(enc, pass); } catch (e) { flow = null; }
  }
  render();
  if (flow) { checkUpdate(false); sendLogs(false); }
}

if (typeof window !== 'undefined' && typeof document !== 'undefined') boot();
