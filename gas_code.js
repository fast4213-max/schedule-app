// ════════════════════════════════════════════════════
// スケジュール調整アプリ - Google Apps Script
// ════════════════════════════════════════════════════
// 【セットアップ】
// 1. Google スプレッドシートを新規作成
// 2. 拡張機能 > Apps Script でこのコードを貼り付け保存
// 3. デプロイ > 新しいデプロイ
//      種類: ウェブアプリ
//      次のユーザーとして実行: 自分
//      アクセス: 全員（匿名ユーザーを含む）
// 4. 発行URLを index.html の GAS_URL に貼る
// ════════════════════════════════════════════════════

const SHEET_RESPONSES    = 'responses';    // 名前,日付,回答,pin_hash,更新日時
const SHEET_PARTICIPANTS = 'participants'; // 名前,pin_hash,登録日時
const SHEET_SETTINGS     = 'settings';    // key,value

// ── シート取得 / 自動生成 ──
function getOrCreate(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── settings シートの読み書き ──
function getSettings() {
  const sheet = getOrCreate(SHEET_SETTINGS, ['key', 'value']);
  const rows  = sheet.getDataRange().getValues();
  const result = {};
  rows.slice(1).forEach(r => { if (r[0]) result[r[0]] = r[1]; });
  // デフォルト値
  if (!result['adminPin']) result['adminPin'] = '0';
  if (!result['eventName']) result['eventName'] = 'スケジュール調整';
  if (!result['closed']) result['closed'] = 'false';
  return result;
}

function setSettings(updates) {
  const sheet = getOrCreate(SHEET_SETTINGS, ['key', 'value']);
  const rows  = sheet.getDataRange().getValues();
  Object.entries(updates).forEach(([key, value]) => {
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        rows[i][1] = value;
        found = true; break;
      }
    }
    if (!found) {
      sheet.appendRow([key, value]);
      rows.push([key, value]);
    }
  });
}

// ════════════════════════════════════════
// GET
// ════════════════════════════════════════
function doGet(e) {
  const action = e.parameter.action;

  // ダッシュボード + 参加者を取得
  if (action === 'get') {
    const year  = parseInt(e.parameter.year,  10);
    const month = parseInt(e.parameter.month, 10);
    return jsonRes({
      data:         getData(year, month),
      participants: getParticipants(),
    });
  }

  // 設定取得
  // month は JavaScript の Date に合わせて 0始まり（1月=0）で返す
  if (action === 'get_config') {
    const s = getSettings();
    // スプレッドシートには 0始まり の month がそのまま保存されている
    const month1 = s['month1'] ? JSON.parse(s['month1']) : null;
    const month2 = s['month2'] ? JSON.parse(s['month2']) : null;

    // ── closed の判定 ──
    // スプレッドシートは TRUE（真偽値）または 'true'/'TRUE'（文字列）どちらもありえる
    const closedVal = s['closed'];
    const closed = closedVal === true
      || String(closedVal).toLowerCase() === 'true';

    return jsonRes({
      config: {
        eventName: s['eventName'] || 'スケジュール調整',
        closed,
        month1,
        month2,
      }
    });
  }

  // 祝日取得（内閣府CSVから）
  if (action === 'get_holidays') {
    return jsonRes({ holidays: fetchHolidays() });
  }

  return jsonRes({ error: 'unknown action' });
}

// ════════════════════════════════════════
// POST
// ════════════════════════════════════════
function doPost(e) {
  const p = JSON.parse(e.postData.contents);

  switch (p.action) {
    case 'verify_pin':   return jsonRes(verifyPin(p.name, p.pin_hash));
    case 'verify_admin': return jsonRes(verifyAdmin(p.pin_hash));
    case 'submit':       return jsonRes(submitResponse(p));
    case 'save_config':  return jsonRes(saveConfigAction(p.config));
    case 'delete_one':   return jsonRes(deleteOne(p.name));
    case 'clear_all':    return jsonRes(clearAll());
    default:             return jsonRes({ status: 'error', message: 'unknown action' });
  }
}

// ── 管理者PIN確認 ──
function verifyAdmin(pinHash) {
  const s = getSettings();
  const stored = s['adminPin'] || '168c00';
  return pinHash === stored ? { status: 'ok' } : { status: 'error' };
}

// ── ユーザーPIN確認 ──
function verifyPin(name, pinHash) {
  const sheet = getOrCreate(SHEET_PARTICIPANTS, ['名前', 'pin_hash', '登録日時']);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === name) {
      return rows[i][1] === pinHash ? { status: 'ok' } : { status: 'pin_error' };
    }
  }
  return { status: 'not_found' };
}

// ── 回答保存 ──
function submitResponse(p) {
  const { name, pin_hash, overwrite, selections } = p;
  const pSheet = getOrCreate(SHEET_PARTICIPANTS, ['名前', 'pin_hash', '登録日時']);
  const pRows  = pSheet.getDataRange().getValues();
  const existingUser = pRows.slice(1).find(r => r[0] === name);

  // 既存ユーザーのPINチェック（上書き時）
  if (existingUser && overwrite && existingUser[1] !== pin_hash) {
    return { status: 'pin_error' };
  }

  const rSheet = getOrCreate(SHEET_RESPONSES, ['名前', '日付', '回答', 'pin_hash', '更新日時']);
  const now = new Date();

  // 上書き：既存の回答行を削除
  if (overwrite) {
    const rRows = rSheet.getDataRange().getValues();
    for (let i = rRows.length - 1; i >= 1; i--) {
      if (rRows[i][0] === name) rSheet.deleteRow(i + 1);
    }
  }

  // 新規回答を追加（日付は必ず文字列で保存）
  selections.forEach(({ date, state }) => {
    const dateStr = String(date).trim();
    rSheet.appendRow([name, dateStr, state, pin_hash, now.toISOString()]);
  });

  // 参加者テーブルに登録
  if (!existingUser) {
    pSheet.appendRow([name, pin_hash, now.toISOString()]);
  }

  return { status: 'ok' };
}

// ── 設定保存 ──
// フロントから month は 0始まりで届くので、そのまま保存する
function saveConfigAction(cfg) {
  const updates = {
    eventName: cfg.eventName || 'スケジュール調整',
    closed:    String(cfg.closed === true || cfg.closed === 'true'),
    month1:    cfg.month1 ? JSON.stringify(cfg.month1) : '',
    month2:    cfg.month2 ? JSON.stringify(cfg.month2) : '',
  };
  if (cfg.adminPin) {
    updates['adminPin'] = cfg.adminPin;
  }
  setSettings(updates);
  return { status: 'ok' };
}

// ── 1人削除 ──
function deleteOne(name) {
  const rSheet = getOrCreate(SHEET_RESPONSES, ['名前', '日付', '回答', 'pin_hash', '更新日時']);
  let rRows = rSheet.getDataRange().getValues();
  for (let i = rRows.length - 1; i >= 1; i--) {
    if (rRows[i][0] === name) rSheet.deleteRow(i + 1);
  }
  const pSheet = getOrCreate(SHEET_PARTICIPANTS, ['名前', 'pin_hash', '登録日時']);
  let pRows = pSheet.getDataRange().getValues();
  for (let i = pRows.length - 1; i >= 1; i--) {
    if (pRows[i][0] === name) pSheet.deleteRow(i + 1);
  }
  return { status: 'ok' };
}

// ── 全員クリア ──
function clearAll() {
  const rSheet = getOrCreate(SHEET_RESPONSES, ['名前', '日付', '回答', 'pin_hash', '更新日時']);
  const rLast = rSheet.getLastRow();
  if (rLast > 1) rSheet.deleteRows(2, rLast - 1);

  const pSheet = getOrCreate(SHEET_PARTICIPANTS, ['名前', 'pin_hash', '登録日時']);
  const pLast = pSheet.getLastRow();
  if (pLast > 1) pSheet.deleteRows(2, pLast - 1);

  return { status: 'ok' };
}

// ── 集計 ──
// month は 1始まり（1月=1）で受け取る
function getData(year, month) {
  const sheet = getOrCreate(SHEET_RESPONSES, ['名前', '日付', '回答', 'pin_hash', '更新日時']);
  const rows  = sheet.getDataRange().getValues();
  const result = {};
  rows.slice(1).forEach(row => {
    const name  = row[0];
    const raw   = row[1];
    const state = row[2];
    if (!raw || !state) return;

    let dateStr;
    if (raw instanceof Date) {
      const y = raw.getFullYear();
      const m = String(raw.getMonth() + 1).padStart(2, '0');
      const d = String(raw.getDate()).padStart(2, '0');
      dateStr = `${y}-${m}-${d}`;
    } else {
      dateStr = String(raw).replace(/\//g, '-').trim();
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        dateStr = `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
      }
    }

    const [y, m] = dateStr.split('-').map(Number);
    if (y !== year || m !== month) return;
    if (!result[dateStr]) result[dateStr] = { ok: [], maybe: [], ng: [] };
    if (!result[dateStr][state].includes(name)) result[dateStr][state].push(name);
  });
  return result;
}

// ── 参加者一覧 ──
function getParticipants() {
  const sheet = getOrCreate(SHEET_PARTICIPANTS, ['名前', 'pin_hash', '登録日時']);
  return sheet.getDataRange().getValues().slice(1).map(r => ({ name: r[0] }));
}

// ── レスポンス ──
function jsonRes(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════
// 祝日取得（内閣府CSV）
// ════════════════════════════════════════
function fetchHolidays() {
  const settings = getSettings();
  const cachedDate = settings['holidays_cache_date'] || '';
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  if (cachedDate === today && settings['holidays_cache']) {
    try {
      return JSON.parse(settings['holidays_cache']);
    } catch(e) {}
  }

  try {
    const res = UrlFetchApp.fetch(
      'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv',
      { muteHttpExceptions: true }
    );
    const text = res.getContentText('Shift_JIS');
    const holidays = parseHolidayCsv(text);

    setSettings({
      holidays_cache: JSON.stringify(holidays),
      holidays_cache_date: today,
    });

    return holidays;
  } catch(e) {
    Logger.log('祝日取得エラー: ' + e);
    if (settings['holidays_cache']) {
      try { return JSON.parse(settings['holidays_cache']); } catch(e2) {}
    }
    return {};
  }
}

function parseHolidayCsv(text) {
  const result = {};
  const lines = text.split(/\r?\n/);
  lines.forEach((line, idx) => {
    if (idx === 0 || !line.trim()) return;
    const cols = line.split(',');
    if (cols.length < 2) return;
    const dateStr = cols[0].trim();
    const name    = cols[1].trim();
    if (!dateStr || !name) return;
    const parts = dateStr.split('/');
    if (parts.length !== 3) return;
    const key = `${parts[0]}-${String(parts[1]).padStart(2,'0')}-${String(parts[2]).padStart(2,'0')}`;
    result[key] = name;
  });
  return result;
}
