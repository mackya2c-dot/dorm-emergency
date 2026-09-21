/**
 * 寮 緊急時対応ガイド（PWA）管理用 GAS
 * - スプレッドシートでフローを編集
 * - 分岐チェック → 暗号化 → GitHub へ公開
 * - 印刷用PDFの作成
 * - PWA からの対応記録（ログ）の受信（doPost）
 */

const SHEET = { SET: '設定', TOP: 'トップ', NODE: 'ノード', CONTACT: '連絡先', LOG: 'ログ' };

const HEAD = {
  SET: ['項目', '値', '説明'],
  TOP: ['並び', 'ボタン名_ja', 'ボタン名_en', '開始ID', '色'],
  NODE: ['ID', '種別', '緊急度', '表示文_ja', '表示文_en', '詳細_ja', '詳細_en',
    '選択肢1_ja', '選択肢1_en', '次ID1', '選択肢2_ja', '選択肢2_en', '次ID2',
    '選択肢3_ja', '選択肢3_en', '次ID3', '選択肢4_ja', '選択肢4_en', '次ID4',
    '連絡先ID', 'やらないこと_ja', 'やらないこと_en'],
  CONTACT: ['ID', '名称_ja', '名称_en', '電話', '対応時間_ja', '対応時間_en'],
  LOG: ['受信日時', '記録日時', 'RS名', '場面', '経路', '最終ID', '実施済み', '未実施', 'メモ', '版', '記録ID']
};

const LEVEL = { '高': 'high', '中': 'mid', '低': 'low' };
const COLOR = { '赤': 'red', '橙': 'orange', '青': 'blue', '緑': 'green' };
const TYPE = { '質問': 'q', '対応': 'a' };

/* ============================== メニュー ============================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('緊急対応フロー')
    .addItem('① 初期設定（シート作成）', 'setupSheets')
    .addItem('② GitHubトークンを登録', 'registerToken')
    .addSeparator()
    .addItem('分岐をチェック', 'checkFlow')
    .addItem('GitHubへ公開', 'publishFlow')
    .addItem('印刷用PDFを作成', 'createPrintPdf')
    .addToUi();
}

/* ============================== 初期設定 ============================== */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  const made = [];

  if (!ss.getSheetByName(SHEET.SET)) {
    const sh = ss.insertSheet(SHEET.SET);
    const rows = [
      ['アプリ名_ja', '寮 緊急時対応ガイド', 'アプリ上部に表示'],
      ['アプリ名_en', 'Dorm Emergency Guide', ''],
      ['GitHubユーザー名', '', 'リポジトリの持ち主（URLの github.com/〇〇 の部分）'],
      ['リポジトリ名', 'dorm-emergency', ''],
      ['ブランチ', 'main', ''],
      ['パスコード', '', '8文字以上。RSがアプリの初回起動時に入力'],
      ['暗号化の反復回数', 10000, '公開が遅い場合は5000に下げる'],
      ['ログ受信URL', '', 'このGASをウェブアプリとしてデプロイしたURL'],
      ['ログ受信キー', Utilities.getUuid(), '自動生成。変更不要'],
      ['現在の版', '', '公開時に自動記入'],
      ['最終公開日時', '', '公開時に自動記入']
    ];
    sh.getRange(1, 1, 1, 3).setValues([HEAD.SET]);
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
    styleHeader_(sh, 3);
    sh.setColumnWidth(1, 160).setColumnWidth(2, 320).setColumnWidth(3, 380);
    made.push(SHEET.SET);
  }

  if (!ss.getSheetByName(SHEET.TOP)) {
    const sh = ss.insertSheet(SHEET.TOP);
    const rows = sampleTop_();
    sh.getRange(1, 1, 1, HEAD.TOP.length).setValues([HEAD.TOP]);
    sh.getRange(2, 1, rows.length, HEAD.TOP.length).setValues(rows);
    styleHeader_(sh, HEAD.TOP.length);
    sh.getRange('E2:E100').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(COLOR), true).build());
    sh.setColumnWidths(2, 2, 220);
    made.push(SHEET.TOP);
  }

  if (!ss.getSheetByName(SHEET.NODE)) {
    const sh = ss.insertSheet(SHEET.NODE);
    const rows = sampleNodes_();
    sh.getRange(1, 1, 1, HEAD.NODE.length).setValues([HEAD.NODE]);
    sh.getRange(2, 1, rows.length, HEAD.NODE.length).setValues(rows);
    styleHeader_(sh, HEAD.NODE.length);
    sh.setFrozenColumns(1);
    sh.getRange('B2:B500').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(TYPE), true).build());
    sh.getRange('C2:C500').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(LEVEL), true).build());
    sh.getRange(2, 1, 500, HEAD.NODE.length).setWrap(true).setVerticalAlignment('top');
    sh.setColumnWidths(4, 4, 260);
    sh.setColumnWidths(8, 12, 130);
    sh.setColumnWidths(21, 2, 220);
    made.push(SHEET.NODE);
  }

  if (!ss.getSheetByName(SHEET.CONTACT)) {
    const sh = ss.insertSheet(SHEET.CONTACT);
    const rows = sampleContacts_();
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange(1, 1, 1, HEAD.CONTACT.length).setValues([HEAD.CONTACT]);
    sh.getRange(2, 1, rows.length, HEAD.CONTACT.length).setValues(rows);
    styleHeader_(sh, HEAD.CONTACT.length);
    sh.setColumnWidths(2, 2, 280);
    made.push(SHEET.CONTACT);
  }

  if (!ss.getSheetByName(SHEET.LOG)) {
    const sh = ss.insertSheet(SHEET.LOG);
    sh.getRange(1, 1, 1, HEAD.LOG.length).setValues([HEAD.LOG]);
    styleHeader_(sh, HEAD.LOG.length);
    sh.setColumnWidth(5, 420).setColumnWidth(9, 280);
    made.push(SHEET.LOG);
  }

  const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  SpreadsheetApp.getUi().alert(made.length
    ? '作成しました：' + made.join('、') + '\n\n次に「設定」シートのGitHubユーザー名とパスコードを入力してください。'
    : 'シートはすべて作成済みです（既存のシートは変更していません）。');
}

function styleHeader_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#14213D').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
}

function registerToken() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('GitHubトークンを登録',
    'github_pat_ で始まるトークンを貼り付けてください。\n（スクリプトプロパティに保存され、シートには表示されません）',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const token = res.getResponseText().trim();
  if (!/^(github_pat_|ghp_)/.test(token)) { ui.alert('トークンの形式が正しくありません。'); return; }
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  ui.alert('登録しました。');
}

/* ============================== 読み込み ============================== */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function readTable_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。先に「① 初期設定」を実行してください。');
  return tableFromValues_(sh.getDataRange().getDisplayValues());
}

/** 2次元配列 → 見出しをキーにしたオブジェクト配列（_row = シートの行番号） */
function tableFromValues_(values) {
  const head = values[0].map(function (h) { return String(h).trim(); });
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.every(function (v) { return String(v).trim() === ''; })) continue;
    const o = { _row: i + 1 };
    head.forEach(function (h, j) { o[h] = String(r[j] == null ? '' : r[j]).trim(); });
    out.push(o);
  }
  return out;
}

function getSettings_() {
  const sh = ss_().getSheetByName(SHEET.SET);
  if (!sh) throw new Error('シート「設定」がありません。');
  const vals = sh.getDataRange().getDisplayValues();
  const s = {};
  for (let i = 1; i < vals.length; i++) s[String(vals[i][0]).trim()] = String(vals[i][1]).trim();
  return s;
}

function setSetting_(key, value) {
  const sh = ss_().getSheetByName(SHEET.SET);
  const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value, '']);
}

function loadAll_() {
  return {
    settings: getSettings_(),
    top: readTable_(SHEET.TOP),
    nodes: readTable_(SHEET.NODE),
    contacts: readTable_(SHEET.CONTACT)
  };
}

/* ============================== 組み立て・チェック ============================== */

/** シートの内容からアプリ用データを組み立て、エラー・警告を返す（副作用なし） */
function buildFlow_(t) {
  const errors = [], warnings = [];
  const E = function (sheet, row, msg) { errors.push('【' + sheet + ' ' + row + '行目】' + msg); };
  const W = function (sheet, row, msg) { warnings.push('【' + sheet + ' ' + row + '行目】' + msg); };
  const tx = function (ja, en) { return { ja: ja || '', en: en || '' }; };

  // 連絡先
  const contacts = {}, contactOrder = [];
  t.contacts.forEach(function (c) {
    const id = c['ID'];
    if (!id) { E('連絡先', c._row, 'IDが空です。'); return; }
    if (contacts[id]) { E('連絡先', c._row, 'ID「' + id + '」が重複しています。'); return; }
    if (!c['名称_ja']) E('連絡先', c._row, '名称_jaが空です。');
    if (!c['電話']) E('連絡先', c._row, '電話が空です。');
    if (!c['名称_en']) W('連絡先', c._row, '名称_enが空です（英語表示では日本語が出ます）。');
    contacts[id] = { name: tx(c['名称_ja'], c['名称_en']), tel: c['電話'], hours: tx(c['対応時間_ja'], c['対応時間_en']) };
    contactOrder.push(id);
  });

  // ノード
  const nodes = {};
  t.nodes.forEach(function (n) {
    const id = n['ID'];
    if (!id) { E('ノード', n._row, 'IDが空です。'); return; }
    if (!/^[A-Za-z0-9_-]+$/.test(id)) E('ノード', n._row, 'ID「' + id + '」には半角英数字・_・- のみ使えます。');
    if (nodes[id]) { E('ノード', n._row, 'ID「' + id + '」が重複しています。'); return; }
    const type = TYPE[n['種別']];
    if (!type) E('ノード', n._row, '種別は「質問」か「対応」にしてください。');
    const level = LEVEL[n['緊急度']] || 'mid';
    if (!LEVEL[n['緊急度']]) W('ノード', n._row, '緊急度が空または不正のため「中」として扱います。');
    if (!n['表示文_ja']) E('ノード', n._row, '表示文_jaが空です。');
    if (!n['表示文_en']) W('ノード', n._row, '表示文_enが空です。');

    const choices = [];
    for (let k = 1; k <= 4; k++) {
      const ja = n['選択肢' + k + '_ja'], en = n['選択肢' + k + '_en'], next = n['次ID' + k];
      if (!ja && !en && !next) continue;
      if (!ja) E('ノード', n._row, '選択肢' + k + '_jaが空です。');
      if (!next) E('ノード', n._row, '選択肢' + k + 'の次IDが空です。');
      if (ja && !en) W('ノード', n._row, '選択肢' + k + '_enが空です。');
      choices.push({ label: tx(ja, en), next: next, _k: k });
    }
    if (type === 'q' && choices.length < 2) E('ノード', n._row, '質問には選択肢が2つ以上必要です。');

    const cids = (n['連絡先ID'] || '').split(/[,、\s]+/).filter(String);
    nodes[id] = {
      id: id, type: type || 'q', level: level,
      text: tx(n['表示文_ja'], n['表示文_en']),
      detail: tx(n['詳細_ja'], n['詳細_en']),
      choices: choices, contacts: cids,
      dont: tx(n['やらないこと_ja'], n['やらないこと_en']),
      _row: n._row
    };
  });

  // 参照チェック
  Object.keys(nodes).forEach(function (id) {
    const n = nodes[id];
    n.choices.forEach(function (c) {
      if (c.next && !nodes[c.next]) E('ノード', n._row, '選択肢' + c._k + 'の次ID「' + c.next + '」が見つかりません。');
      if (c.next === id) W('ノード', n._row, '選択肢' + c._k + 'が自分自身を指しています。');
    });
    n.contacts.forEach(function (cid) {
      if (!contacts[cid]) E('ノード', n._row, '連絡先ID「' + cid + '」が連絡先シートにありません。');
    });
    const dj = n.detail.ja.split('\n').filter(String).length, de = n.detail.en.split('\n').filter(String).length;
    if (n.detail.en && dj !== de) W('ノード', n._row, '詳細の行数が日本語（' + dj + '行）と英語（' + de + '行）で違います。');
  });

  // トップ
  const top = [];
  t.top.slice().sort(function (a, b) { return Number(a['並び'] || 0) - Number(b['並び'] || 0); }).forEach(function (r) {
    if (!r['ボタン名_ja']) E('トップ', r._row, 'ボタン名_jaが空です。');
    if (!r['開始ID'] || !nodes[r['開始ID']]) E('トップ', r._row, '開始ID「' + r['開始ID'] + '」がノードシートにありません。');
    if (!r['ボタン名_en']) W('トップ', r._row, 'ボタン名_enが空です。');
    top.push({ label: tx(r['ボタン名_ja'], r['ボタン名_en']), start: r['開始ID'], color: COLOR[r['色']] || 'blue' });
  });
  if (!top.length) errors.push('【トップ】ボタンが1つもありません。');

  // 到達できないノード
  const seen = {}, queue = top.map(function (x) { return x.start; }).filter(function (id) { return nodes[id]; });
  while (queue.length) {
    const id = queue.shift();
    if (seen[id]) continue;
    seen[id] = true;
    nodes[id].choices.forEach(function (c) { if (nodes[c.next] && !seen[c.next]) queue.push(c.next); });
  }
  Object.keys(nodes).forEach(function (id) {
    if (!seen[id]) W('ノード', nodes[id]._row, 'ID「' + id + '」はトップのどのボタンからもたどれません。');
  });

  // 出力用に内部項目を除去
  Object.keys(nodes).forEach(function (id) {
    delete nodes[id]._row;
    nodes[id].choices.forEach(function (c) { delete c._k; });
  });

  const s = t.settings || {};
  return {
    errors: errors, warnings: warnings,
    payload: {
      meta: { title: tx(s['アプリ名_ja'] || '緊急時対応ガイド', s['アプリ名_en'] || 'Emergency Guide'), version: '', updatedAt: '', logUrl: s['ログ受信URL'] || '', logKey: s['ログ受信キー'] || '' },
      top: top, nodes: nodes, contacts: contacts, contactOrder: contactOrder
    }
  };
}

function checkFlow() {
  const r = buildFlow_(loadAll_());
  showReport_('分岐チェックの結果', r.errors, r.warnings,
    r.errors.length ? '' : 'エラーはありません。公開できます。');
}

function showReport_(title, errors, warnings, okMsg) {
  const esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  let html = '<div style="font-family:sans-serif;font-size:13px;line-height:1.6">';
  if (okMsg) html += '<p style="color:#00875A;font-weight:bold">' + esc(okMsg) + '</p>';
  if (errors.length) html += '<h3 style="color:#C8102E;margin:8px 0">エラー（' + errors.length + '件・公開できません）</h3><ul>' + errors.map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('') + '</ul>';
  if (warnings.length) html += '<h3 style="color:#B45F00;margin:8px 0">警告（' + warnings.length + '件・公開は可能）</h3><ul>' + warnings.map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('') + '</ul>';
  html += '</div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(640).setHeight(480), title);
}

/* ============================== 公開 ============================== */

function publishFlow() {
  const ui = SpreadsheetApp.getUi();
  const t = loadAll_();
  const s = t.settings;
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const pre = [];
  if (!token) pre.push('GitHubトークンが未登録です（メニュー「② GitHubトークンを登録」）。');
  if (!s['GitHubユーザー名']) pre.push('設定シートの「GitHubユーザー名」が空です。');
  if (!s['リポジトリ名']) pre.push('設定シートの「リポジトリ名」が空です。');
  if ((s['パスコード'] || '').length < 8) pre.push('設定シートの「パスコード」を8文字以上で入力してください。');
  if (pre.length) { showReport_('公開できません', pre, [], ''); return; }

  const r = buildFlow_(t);
  if (r.errors.length) { showReport_('公開できません', r.errors, r.warnings, ''); return; }
  if (r.warnings.length) {
    const ok = ui.alert('警告が' + r.warnings.length + '件あります',
      r.warnings.slice(0, 10).join('\n') + (r.warnings.length > 10 ? '\n…ほか' + (r.warnings.length - 10) + '件' : '') + '\n\nこのまま公開しますか？',
      ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
  }

  const now = new Date();
  const version = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  const updatedAt = Utilities.formatDate(now, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
  r.payload.meta.version = version;
  r.payload.meta.updatedAt = updatedAt;

  const iter = Math.max(1000, parseInt(s['暗号化の反復回数'], 10) || 10000);
  const enc = encryptPayload_(JSON.stringify(r.payload), s['パスコード'], iter);

  const gh = { owner: s['GitHubユーザー名'], repo: s['リポジトリ名'], branch: s['ブランチ'] || 'main', token: token };
  // データ本体を先に、版情報を後に書き込む（版情報だけ先に更新されるのを防ぐ）
  githubPut_(gh, 'data/flow.enc.json', JSON.stringify(enc), '緊急対応フロー データ更新 ' + version);
  githubPut_(gh, 'data/version.json', JSON.stringify({ version: version, updatedAt: updatedAt }), '緊急対応フロー 版情報 ' + version);

  setSetting_('現在の版', version);
  setSetting_('最終公開日時', Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  ui.alert('公開しました（版 ' + version + '）。\n\nGitHub Pagesへの反映には1〜数分かかります。\nRSの端末には、次に通信できる状態でアプリを開いたときに反映されます。');
}

function githubPut_(gh, path, content, message) {
  const url = 'https://api.github.com/repos/' + encodeURIComponent(gh.owner) + '/' + encodeURIComponent(gh.repo) + '/contents/' + path;
  const headers = {
    Authorization: 'Bearer ' + gh.token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
  let sha = null;
  const get = UrlFetchApp.fetch(url + '?ref=' + encodeURIComponent(gh.branch), { headers: headers, muteHttpExceptions: true });
  if (get.getResponseCode() === 200) sha = JSON.parse(get.getContentText()).sha;
  else if (get.getResponseCode() !== 404) throw new Error(githubError_(get, path));

  const body = { message: message, content: Utilities.base64Encode(content, Utilities.Charset.UTF_8), branch: gh.branch };
  if (sha) body.sha = sha;
  const put = UrlFetchApp.fetch(url, {
    method: 'put', headers: headers, contentType: 'application/json',
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  const code = put.getResponseCode();
  if (code !== 200 && code !== 201) throw new Error(githubError_(put, path));
}

function githubError_(res, path) {
  const code = res.getResponseCode();
  let msg = '';
  try { msg = JSON.parse(res.getContentText()).message; } catch (e) { msg = res.getContentText().slice(0, 200); }
  const hint = {
    401: 'トークンが無効か期限切れです。作り直して「② GitHubトークンを登録」をやり直してください。',
    403: 'トークンの権限が不足しています（Contents: Read and write が必要）。',
    404: 'リポジトリが見つかりません。GitHubユーザー名・リポジトリ名と、トークンの対象リポジトリを確認してください。',
    409: '書き込みが競合しました。少し待ってからもう一度公開してください。',
    422: 'ブランチ名などの指定が正しくありません。'
  }[code] || '';
  return 'GitHubへの書き込みに失敗しました（' + path + '、HTTP ' + code + '）。' + hint + '\n詳細：' + msg;
}

/* ============================== 暗号化 ==============================
 * 鍵導出：PBKDF2-HMAC-SHA256（64バイト → 前半を暗号用、後半を改ざん検知用）
 * 暗号化：HMAC-SHA256 をカウンタモードで使った鍵ストリームとのXOR
 * 改ざん検知：HMAC-SHA256(salt || nonce || 暗号文)
 * ブラウザ側は Web Crypto API で同じ処理を行う（app.js の decryptFlow）
 * ================================================================== */

function encryptPayload_(plainText, passcode, iter) {
  const pass = utf8Bytes_(passcode);
  const salt = randomBytes_(16);
  const nonce = randomBytes_(16);
  const dk = pbkdf2_(pass, salt, iter, 2);
  const encKey = dk.slice(0, 32), macKey = dk.slice(32, 64);
  const pt = utf8Bytes_(plainText);
  const ct = new Array(pt.length);
  const blocks = Math.ceil(pt.length / 32);
  for (let b = 0; b < blocks; b++) {
    const ks = hmac_(nonce.concat(int32be_(b)), encKey);
    for (let j = 0; j < 32; j++) {
      const i = b * 32 + j;
      if (i >= pt.length) break;
      ct[i] = sb_(pt[i] ^ ks[j]);
    }
  }
  const mac = hmac_(salt.concat(nonce, ct), macKey);
  return {
    v: 1, alg: 'PBKDF2-HMAC-SHA256/HMAC-SHA256-CTR/HMAC-SHA256', iter: iter,
    salt: Utilities.base64Encode(salt), nonce: Utilities.base64Encode(nonce),
    ct: Utilities.base64Encode(ct), mac: Utilities.base64Encode(mac)
  };
}

function pbkdf2_(pass, salt, iter, blocks) {
  let out = [];
  for (let b = 1; b <= blocks; b++) {
    let u = hmac_(salt.concat(int32be_(b)), pass);
    const t = u.slice();
    for (let i = 1; i < iter; i++) {
      u = hmac_(u, pass);
      for (let j = 0; j < 32; j++) t[j] = sb_(t[j] ^ u[j]);
    }
    out = out.concat(t);
  }
  return out;
}

function hmac_(value, key) {
  return Array.prototype.slice.call(Utilities.computeHmacSha256Signature(value, key));
}

function utf8Bytes_(s) {
  return Array.prototype.slice.call(Utilities.newBlob('').setDataFromString(String(s), 'UTF-8').getBytes());
}

function randomBytes_(n) {
  const out = [];
  while (out.length < n) {
    const hex = Utilities.getUuid().replace(/-/g, '');
    for (let i = 0; i < 32; i += 2) out.push(sb_(parseInt(hex.substr(i, 2), 16)));
  }
  return out.slice(0, n);
}

function int32be_(i) {
  return [(i >>> 24) & 255, (i >>> 16) & 255, (i >>> 8) & 255, i & 255].map(sb_);
}

/** 0〜255 または任意の整数 → GASのバイト（-128〜127） */
function sb_(v) { return ((v & 255) << 24) >> 24; }

/* ============================== 印刷用PDF ============================== */

function createPrintPdf() {
  const t = loadAll_();
  const r = buildFlow_(t);
  if (r.errors.length) { showReport_('PDFを作成できません', r.errors, r.warnings, ''); return; }
  const html = buildPrintHtml_(r.payload);
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm');
  const blob = Utilities.newBlob(html, 'text/html', 'print.html').getAs('application/pdf')
    .setName('緊急時対応フロー_印刷用_' + stamp + '.pdf');
  const parents = DriveApp.getFileById(ss_().getId()).getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const file = folder.createFile(blob);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<p style="font-family:sans-serif">作成しました：<br><a href="' + file.getUrl() + '" target="_blank">' + file.getName() + '</a></p>').setWidth(460).setHeight(140),
    '印刷用PDF');
}

function buildPrintHtml_(p) {
  const esc = function (s) { return String(s || '').replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  const lines = function (s) { return String(s || '').split('\n').map(function (x) { return x.trim(); }).filter(String); };
  const lvColor = { high: '#C8102E', mid: '#B45F00', low: '#1F5FAD' };

  let h = '<html><head><meta charset="UTF-8"><style>' +
    'body{font-family:sans-serif;font-size:10.5pt;color:#14213D;line-height:1.5}' +
    'h1{font-size:18pt;margin:0 0 4pt}h2{font-size:14pt;margin:0 0 8pt;padding:4pt 8pt;color:#fff}' +
    '.en{color:#555;font-size:9pt}.node{border:1pt solid #999;margin:0 0 8pt;padding:6pt 8pt;page-break-inside:avoid}' +
    '.id{display:inline-block;font-weight:bold;color:#fff;padding:1pt 6pt;margin-right:6pt}' +
    '.q{font-weight:bold;font-size:11.5pt}table{border-collapse:collapse;width:100%}' +
    'td,th{border:1pt solid #999;padding:3pt 6pt;vertical-align:top;text-align:left}' +
    '.dont{border-left:3pt solid #C8102E;padding-left:6pt;margin-top:4pt}.brk{page-break-before:always}' +
    '</style></head><body>';

  h += '<h1>' + esc(p.meta.title.ja) + '</h1><div class="en">' + esc(p.meta.title.en) + '</div>';
  h += '<p>作成日：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日') + '　／　通信やアプリが使えないときは、この紙の番号（ID）をたどってください。</p>';
  h += '<h2 style="background:#14213D">連絡先 / Contacts</h2><table><tr><th>名称</th><th>電話</th><th>対応時間</th></tr>';
  p.contactOrder.forEach(function (id) {
    const c = p.contacts[id];
    h += '<tr><td>' + esc(c.name.ja) + '<div class="en">' + esc(c.name.en) + '</div></td><td style="font-size:13pt;font-weight:bold">' + esc(c.tel) + '</td><td>' + esc(c.hours.ja) + '</td></tr>';
  });
  h += '</table>';

  const colorHex = { red: '#C8102E', orange: '#B45F00', blue: '#1F5FAD', green: '#00875A' };
  p.top.forEach(function (tp) {
    h += '<div class="brk"></div><h2 style="background:' + (colorHex[tp.color] || '#1F5FAD') + '">' + esc(tp.label.ja) + ' / ' + esc(tp.label.en) + '　（' + esc(tp.start) + ' から開始）</h2>';
    const order = [], seen = {}, q = [tp.start];
    while (q.length) {
      const id = q.shift();
      if (seen[id] || !p.nodes[id]) continue;
      seen[id] = true; order.push(id);
      p.nodes[id].choices.forEach(function (c) { q.push(c.next); });
    }
    order.forEach(function (id) {
      const n = p.nodes[id];
      h += '<div class="node"><span class="id" style="background:' + lvColor[n.level] + '">' + esc(id) + '</span>';
      h += '<span class="q">' + esc(n.text.ja) + '</span><div class="en">' + esc(n.text.en) + '</div>';
      const dj = lines(n.detail.ja), de = lines(n.detail.en);
      if (dj.length) {
        h += '<ul style="margin:4pt 0">';
        dj.forEach(function (l, i) { h += '<li>' + (n.type === 'a' ? '□ ' : '') + esc(l) + (de[i] ? '<div class="en">' + esc(de[i]) + '</div>' : '') + '</li>'; });
        h += '</ul>';
      }
      if (n.contacts.length) {
        h += '<div>連絡先：' + n.contacts.map(function (cid) { const c = p.contacts[cid]; return c ? esc(c.name.ja) + ' <b>' + esc(c.tel) + '</b>' : ''; }).join('　') + '</div>';
      }
      if (n.dont.ja) h += '<div class="dont"><b>やらないこと：</b>' + esc(n.dont.ja) + '<div class="en">Do not: ' + esc(n.dont.en) + '</div></div>';
      if (n.choices.length) {
        h += '<table style="margin-top:4pt">';
        n.choices.forEach(function (c) { h += '<tr><td style="width:75%">' + esc(c.label.ja) + ' <span class="en">/ ' + esc(c.label.en) + '</span></td><td><b>→ ' + esc(c.next) + '</b></td></tr>'; });
        h += '</table>';
      } else {
        h += '<div style="margin-top:4pt"><b>ここで終了。対応内容を記録する。</b></div>';
      }
      h += '</div>';
    });
  });
  return h + '</body></html>';
}

/* ============================== ログ受信（ウェブアプリ） ============================== */

function doGet() {
  return ContentService.createTextOutput('OK');
}

function doPost(e) {
  const out = function (o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); };
  try {
    const body = JSON.parse(e.postData.contents);
    const s = getSettings_();
    if (!body || !s['ログ受信キー'] || body.key !== s['ログ受信キー']) return out({ ok: false, error: 'key' });
    const logs = Array.isArray(body.logs) ? body.logs.slice(0, 200) : [];
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = ss_().getSheetByName(SHEET.LOG);
      const last = sh.getLastRow();
      const idCol = HEAD.LOG.length;
      const existing = {};
      if (last > 1) sh.getRange(2, idCol, last - 1, 1).getValues().forEach(function (r) { existing[r[0]] = true; });
      const now = new Date();
      const rows = [];
      logs.forEach(function (l) {
        const id = String(l.id || '').slice(0, 64);
        if (!id || existing[id]) return;
        existing[id] = true;
        const clip = function (v, n) { return String(v == null ? '' : v).slice(0, n || 2000); };
        rows.push([now, l.at ? new Date(l.at) : '', clip(l.rs, 100), clip(l.scenario, 200), clip(l.path),
          clip(l.end, 64), clip(l.done), clip(l.notDone), clip(l.memo), clip(l.version, 40), id]);
      });
      if (rows.length) sh.getRange(last + 1, 1, rows.length, HEAD.LOG.length).setValues(rows);
      return out({ ok: true, received: rows.length });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

/* ============================== サンプルデータ ============================== */

function sampleTop_() {
  return [
    [1, '地震', 'Earthquake', 'EQ1', '赤'],
    [2, '急病・けが', 'Sudden illness / injury', 'MED1', '赤'],
    [3, '火災報知器が鳴った', 'Fire alarm is sounding', 'FA1', '橙'],
    [4, '水漏れ・設備の故障', 'Water leak / facility trouble', 'WT1', '青'],
    [5, '連絡がとれない入居者', 'Resident cannot be reached', 'MS1', '青']
  ];
}

function sampleContacts_() {
  return [
    ['FIRE', '消防・救急', 'Fire / Ambulance', '119', '24時間', '24 hours'],
    ['POLICE', '警察', 'Police', '110', '24時間', '24 hours'],
    ['Q7119', '救急相談センター（東京）', 'Emergency Advice Line (Tokyo)', '#7119', '24時間', '24 hours'],
    ['UNIV', '大学 夜間・休日緊急連絡先（要設定）', 'University after-hours emergency line (to be set)', '000-0000-0000', '24時間', '24 hours'],
    ['KANRI', '寮管理会社 24時間窓口（要設定）', 'Dorm management 24h desk (to be set)', '000-0000-0001', '24時間', '24 hours'],
    ['KEIBI', '警備会社（要設定）', 'Security company (to be set)', '000-0000-0002', '24時間', '24 hours'],
    ['D171', '災害用伝言ダイヤル', 'Disaster Emergency Message Dial', '171', '災害時', 'During disasters']
  ];
}

function sampleNodes_() {
  // [ID, 種別, 緊急度, 表示文ja, 表示文en, 詳細ja[], 詳細en[], 選択肢[[ja,en,次ID]], 連絡先ID, やらないことja, やらないことen]
  const N = function (id, type, lv, tja, ten, dja, den, ch, con, xja, xen) {
    const row = [id, type, lv, tja, ten, (dja || []).join('\n'), (den || []).join('\n')];
    for (let k = 0; k < 4; k++) { const c = (ch || [])[k]; row.push(c ? c[0] : '', c ? c[1] : '', c ? c[2] : ''); }
    row.push(con || '', xja || '', xen || '');
    return row;
  };
  return [
    // 地震
    N('EQ1', '質問', '高', '揺れは収まりましたか？', 'Has the shaking stopped?', [], [],
      [['収まった', 'Yes, it has stopped', 'EQ2'], ['まだ揺れている', 'No, still shaking', 'EQA0']]),
    N('EQA0', '対応', '高', 'まず自分の身を守る', 'Protect yourself first',
      ['頭を守り、丈夫な机の下などに入る', '窓・棚・照明から離れる', '揺れている間は外に飛び出さない'],
      ['Protect your head and get under a sturdy desk', 'Stay away from windows, shelves and lights', 'Do not run outside while it is shaking'],
      [['揺れが収まった', 'The shaking has stopped', 'EQ2']]),
    N('EQ2', '質問', '高', '火災やガスの臭いはありますか？', 'Is there fire or a smell of gas?',
      ['自分のいる場所と廊下を確認する'], ['Check where you are and the corridor'],
      [['ある', 'Yes', 'EQA1'], ['ない', 'No', 'EQ3'], ['わからない', 'Not sure', 'EQA2']]),
    N('EQA2', '対応', '中', '安全な範囲で確認する', 'Check only where it is safe',
      ['廊下で煙・焦げ臭さ・ガス臭を確認する', '閉まっているドアは手の甲で熱さを確かめてから開ける', '一人で危険な場所に入らない'],
      ['Check the corridor for smoke, burning or gas smells', 'Touch closed doors with the back of your hand before opening', 'Do not enter dangerous areas alone'],
      [['火災・ガスの臭いがあった', 'Found fire or gas', 'EQA1'], ['なかった', 'Found nothing', 'EQ3']]),
    N('EQA1', '対応', '高', '火災・ガスへの対応', 'Respond to fire or gas',
      ['119番に通報する（住所は掲示を読み上げる）', '「火事だ」「ガスだ」と大声で周囲に知らせる', '初期消火は炎が天井に届く前まで。無理なら避難する', 'ガスの場合は窓を開け、火気や電気のスイッチに触れない', '階段で避難する（エレベーターは使わない）'],
      ['Call 119 (read out the address on the notice)', 'Shout to alert others: "Fire!" or "Gas!"', 'Fight the fire only before flames reach the ceiling; otherwise evacuate', 'For gas, open windows and do not touch flames or electric switches', 'Evacuate by the stairs (do not use elevators)'],
      [['避難した', 'Evacuated', 'EQA3']], 'FIRE, UNIV',
      '煙の中や、荷物を取りに部屋へ戻らない', 'Do not go back into smoke or to your room for belongings'),
    N('EQ3', '質問', '高', '建物に大きな被害はありますか？', 'Is there major damage to the building?',
      ['壁の大きな亀裂、傾き、ドアが開かない、天井の落下など'], ['Large cracks, tilting, jammed doors, fallen ceilings, etc.'],
      [['ある', 'Yes', 'EQA3'], ['ない', 'No', 'EQA4'], ['わからない', 'Not sure', 'EQA3']]),
    N('EQA3', '対応', '高', '全員を避難場所へ誘導する', 'Lead everyone to the evacuation site',
      ['避難場所：〇〇公園（要設定）', '各階に声をかけて避難を促す', '点呼表・ペン・懐中電灯を持つ', '階段を使う（エレベーターは使わない）', '避難場所で点呼をとる'],
      ['Evacuation site: XX Park (to be set)', 'Call out on each floor to evacuate', 'Take the roll-call sheet, a pen and a flashlight', 'Use the stairs (do not use elevators)', 'Take roll call at the evacuation site'],
      [['点呼が終わった', 'Roll call is done', 'EQ5']], '',
      '建物に戻らない', 'Do not go back into the building'),
    N('EQA4', '対応', '中', '建物内で点呼をとる', 'Take roll call inside the building',
      ['各階に声をかけ、全員の無事を確認する', '点呼表に記入する', '余震に備え、倒れやすい物から離れるよう伝える'],
      ['Check on every floor that everyone is safe', 'Fill in the roll-call sheet', 'Tell residents to stay away from things that could fall in aftershocks'],
      [['点呼が終わった', 'Roll call is done', 'EQ5']]),
    N('EQ5', '質問', '高', 'けが人や、所在がわからない入居者はいますか？', 'Is anyone injured or unaccounted for?', [], [],
      [['けが人がいる', 'Someone is injured', 'MED1'], ['所在がわからない人がいる', 'Someone is unaccounted for', 'MSA1'], ['いない', 'No', 'EQA5']]),
    N('EQA5', '対応', '中', '大学へ状況を報告する', 'Report to the university',
      ['棟名・人数・けが人・建物の状況を報告する', '電話がつながらない場合はSMSや災害用伝言ダイヤル171を使う', '余震に注意し、テレビ・ラジオで情報を確認する'],
      ['Report the building, number of people, injuries and building condition', 'If calls do not connect, use SMS or the Disaster Message Dial 171', 'Watch for aftershocks and follow TV or radio updates'],
      [], 'UNIV, D171'),

    // 急病・けが
    N('MED1', '質問', '高', '呼びかけに反応しますか？', 'Does the person respond when you call out?',
      ['肩を軽くたたきながら大きな声で呼びかける'], ['Tap their shoulder and speak loudly'],
      [['反応しない', 'No response', 'MEDA1'], ['反応する', 'Responds', 'MED2'], ['判断できない', 'Cannot tell', 'MEDA1']]),
    N('MEDA1', '対応', '高', '119番通報とAED', 'Call 119 and get an AED',
      ['119番に通報する（スピーカーにして指示を聞く）', '周りの人を呼び、AEDを持ってきてもらう', '普段どおりの呼吸がなければ胸骨圧迫を始める', '救急隊が着いたら部屋まで案内する', '大学へ連絡する'],
      ['Call 119 (use the speaker and follow instructions)', 'Call others for help and have someone bring the AED', 'If breathing is not normal, start chest compressions', 'Guide the ambulance crew to the room', 'Contact the university'],
      [], 'FIRE, UNIV',
      '自分で病院へ運ばない', 'Do not take the person to hospital yourself'),
    N('MED2', '質問', '高', '次のどれかに当てはまりますか？', 'Do any of these apply?',
      ['呼吸が苦しい', '激しい痛み・大量の出血', 'けいれんしている', 'ろれつが回らない・手足が動かない'],
      ['Difficulty breathing', 'Severe pain or heavy bleeding', 'Having a seizure', 'Slurred speech or weakness in the arms or legs'],
      [['当てはまる', 'Yes', 'MEDA1'], ['当てはまらない', 'No', 'MED3'], ['わからない', 'Not sure', 'MEDA2']]),
    N('MEDA2', '対応', '中', '救急相談（#7119）に電話する', 'Call the emergency advice line (#7119)',
      ['#7119に電話して症状を伝える', '救急車が必要と言われたら119番に通報する', '大学へ報告する'],
      ['Call #7119 and describe the symptoms', 'If they say an ambulance is needed, call 119', 'Report to the university'],
      [], 'Q7119, FIRE, UNIV',
      '薬を渡さない', 'Do not give any medicine'),
    N('MED3', '質問', '中', '本人は自分で病院に行けますか？', 'Can the person go to a clinic on their own?', [], [],
      [['行ける', 'Yes', 'MEDA3'], ['行けない', 'No', 'MEDA2']]),
    N('MEDA3', '対応', '低', '受診を案内する', 'Help them get medical care',
      ['在留カードと健康保険証を持って受診するよう伝える', '夜間なら#7119で受診先を確認する', '翌朝、大学へ報告する'],
      ['Remind them to bring their residence card and health insurance card', 'At night, ask #7119 where to go', 'Report to the university the next morning'],
      [], 'Q7119',
      '薬を渡さない', 'Do not give any medicine'),

    // 火災報知器
    N('FA1', '質問', '高', '煙や炎が見えますか？', 'Can you see smoke or flames?',
      ['受信機（表示盤）で発報場所を確認する'], ['Check the fire alarm panel for the location'],
      [['見える', 'Yes', 'FAA1'], ['見えない', 'No', 'FAA2'], ['確認できない', 'Cannot check', 'FAA1']]),
    N('FAA1', '対応', '高', '火災として対応する', 'Treat it as a fire',
      ['119番に通報する', '「火事だ」と大声で知らせ、避難を呼びかける', '階段で建物の外へ避難する', '避難後に点呼をとる', '大学と管理会社へ連絡する'],
      ['Call 119', 'Shout "Fire!" and tell everyone to evacuate', 'Evacuate outside by the stairs', 'Take roll call after evacuating', 'Contact the university and the management company'],
      [], 'FIRE, UNIV, KANRI',
      '煙の中に戻らない', 'Do not go back into smoke'),
    N('FAA2', '対応', '中', '発報場所を確認する', 'Check the alarm location',
      ['2人以上で発報場所へ向かう', '部屋の場合は外から声をかける', '誤作動や調理の煙だとわかったら管理会社へ連絡する'],
      ['Go to the location with at least one other person', 'If it is a room, call out from outside', 'If it is a false alarm or cooking smoke, contact the management company'],
      [['煙や炎を見つけた', 'Found smoke or flames', 'FAA1']], 'KANRI, KEIBI',
      '受信機を自分の判断で操作・復旧しない', 'Do not operate or reset the alarm panel on your own'),

    // 水漏れ・設備
    N('WT1', '質問', '中', '水が止まらない、または階下に漏れていますか？', 'Is water still running or leaking downstairs?', [], [],
      [['はい', 'Yes', 'WTA1'], ['いいえ', 'No', 'WTA2']]),
    N('WTA1', '対応', '中', '被害の拡大を防ぐ', 'Limit the damage',
      ['管理会社の24時間窓口に連絡する', '止水栓の場所がわかれば閉める', '階下の部屋に声をかける', '電気製品を水から離す'],
      ["Call the management company's 24-hour desk", 'Close the water valve if you know where it is', 'Alert the room below', 'Move electrical items away from the water'],
      [], 'KANRI',
      '濡れたコンセントやスイッチに触らない', 'Do not touch wet outlets or switches'),
    N('WTA2', '対応', '低', '記録して翌日に報告する', 'Record it and report the next day',
      ['場所と状況の写真を撮る', '翌営業日に大学へ報告する'],
      ['Take photos of the location and situation', 'Report to the university on the next business day'],
      []),

    // 連絡がとれない入居者
    N('MS1', '質問', '中', '次のどれかに当てはまりますか？', 'Do any of these apply?',
      ['24時間以上、誰も姿を見ていない', '体調不良や、心配な言動があった', '部屋から異臭・異音がする'],
      ['No one has seen them for over 24 hours', 'They were unwell or said or did something worrying', 'An unusual smell or sound is coming from the room'],
      [['当てはまる', 'Yes', 'MSA1'], ['当てはまらない', 'No', 'MSA2']]),
    N('MSA1', '対応', '高', 'すぐ大学へ連絡する', 'Contact the university now',
      ['夜間でも大学の緊急連絡先へ電話する', '部屋の前から声をかけ、ノックする', '入室するかどうかは大学・管理会社が判断する', '命に関わる状況が疑われる場合は119番に通報する'],
      ['Call the university emergency line, even at night', 'Knock and call out from outside the room', 'The university and management company decide whether to enter', 'If a life may be at risk, call 119'],
      [], 'UNIV, KANRI, FIRE',
      '合鍵などを使って自分から入室しない', 'Do not enter the room yourself (e.g., with a spare key)'),
    N('MSA2', '対応', '低', '本人と連絡をとる', 'Try to reach the resident',
      ['SNSや電話で本人に連絡する', '親しい入居者に最後に会った日時を聞く', '翌朝までに連絡がとれなければ大学へ報告する'],
      ['Message or call the resident', 'Ask close friends when they last saw them', 'If still unreachable by the next morning, report to the university'],
      [], 'UNIV')
  ];
}
