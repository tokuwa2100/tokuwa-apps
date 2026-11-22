/**
 * 勤怠の受け口（TOKUWAボード → Googleドライブ「勤怠」フォルダ）
 *
 * ・ボードの個人日報の画面で「出勤簿をドライブに保存」を押すと、ここに届きます。
 * ・マイドライブに「勤怠」フォルダを作り、その中に「2026年11月」のような月のフォルダを作って、
 *   出勤簿（スプレッドシート：出勤簿・明細・休みの3枚）と、元データ（.json）を保存します。
 * ・同じ月をもう一度保存すると、上書きされます。
 * ・ボードにログインしている人（下の ALLOWED のアドレス）からしか受け付けません。
 */
const API_KEY = 'AIzaSyB5r7pjt6kHVzQ7NKiQ1wBQxp74j_I4wt0';   // Firebase（tokuwa-apps）の鍵。ログインの確認だけに使う
const ALLOWED = ['board@tokuwa.local', 'tokuwa2100@if-n.ne.jp'];
const ROOT_NAME = '勤怠';
const VERSION = '1.2.0';

/* ── 現場日報（コア日報・常用日報）を、個人日報の下書きに使う ──
 * 日報アプリが同じドライブに作っている「日報台帳」（スプレッドシート）を、その日の分だけ直接読みます。
 * 合言葉はいりません（この受け口は、ボードにログインしている人からしか受け付けないため）。 */
const NIPPO_SHEET = '日報台帳';
const NIPPO_COL = { id: 1, date: 3, deleted: 10, dataId: 12, data: 14 };   // 台帳の列（recordId・作業日・削除・データID・データ）

/* 最初に1回、上のメニューでこの関数を選んで「実行」→ 許可すると「勤怠」フォルダができます */
function setup() {
  Logger.log('勤怠フォルダ：' + rootFolder_().getUrl());
}

function doGet() {
  return out_({ ok: true, app: '勤怠の受け口' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const email = verify_(req.idToken);
    if (!email) return out_({ ok: false, error: 'ログインを確認できませんでした' });
    if (req.action === 'ping') return out_({ ok: true, email: email, folder: rootFolder_().getUrl(), version: VERSION });
    if (req.action === 'saveMonth') return out_(saveMonth_(req, email));
    if (req.action === 'siteDay') return out_(siteDay_(req));
    return out_({ ok: false, error: '知らない操作です' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/* ボードのログイン（Firebase）が本物かを確かめて、メールアドレスを返す */
function verify_(idToken) {
  if (!idToken) return null;
  const r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + API_KEY, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) return null;
  const u = (JSON.parse(r.getContentText()).users || [])[0];
  return u && ALLOWED.indexOf(u.email) >= 0 ? u.email : null;
}

function rootFolder_() {
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(ROOT_NAME);
  return it.hasNext() ? it.next() : root.createFolder(ROOT_NAME);
}
function sub_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function saveMonth_(req, email) {
  const m = String(req.month || '');                         // 例 "2026-11"
  if (!/^\d{4}-\d{2}$/.test(m)) return { ok: false, error: '月の指定が正しくありません' };
  const y = m.slice(0, 4), mo = Number(m.slice(5, 7));
  const folder = sub_(rootFolder_(), y + '年' + mo + '月');

  const name = '出勤簿_' + m;
  let ss;
  const it = folder.getFilesByName(name);
  if (it.hasNext()) ss = SpreadsheetApp.open(it.next());
  else { ss = SpreadsheetApp.create(name); DriveApp.getFileById(ss.getId()).moveTo(folder); }
  writeSheet_(ss, '出勤簿', req.summary);
  writeSheet_(ss, '明細', req.detail);
  writeSheet_(ss, '休み', req.leaves);
  ['シート1', 'Sheet1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ss.getSheetByName('出勤簿').getRange('A1').setNote('保存：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '（' + email + '）');

  const jn = '個人日報_' + m + '.json';
  const body = JSON.stringify(req.raw || {}, null, 1);
  const jt = folder.getFilesByName(jn);
  if (jt.hasNext()) jt.next().setContent(body); else folder.createFile(jn, body, 'application/json');

  return { ok: true, url: ss.getUrl(), folder: folder.getUrl() };
}

function writeSheet_(ss, title, rows) {
  rows = rows || [];
  const sh = ss.getSheetByName(title) || ss.insertSheet(title);
  sh.clear();
  if (!rows.length) return;
  const w = Math.max.apply(null, rows.map(function (r) { return r.length; }));
  const data = rows.map(function (r) { const a = r.slice(); while (a.length < w) a.push(''); return a; });
  sh.getRange(1, 1, data.length, w).setValues(data);
  sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#e3f3e7');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, w);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════ 現場日報 → 個人日報の下書き ═══════════
 * ボードで作業員が日報を書く日を選ぶと、その日・その人の現場日報だけを、その場で読んで返します。
 * 読んで返すだけで、どこにも保存しません。 */
function siteDay_(req) {
  const date = String(req.date || ''), name = String(req.name || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name) return { ok: false, error: '日付か名前がありません' };
  return { ok: true, date: date, records: readDay_(date, name) };
}

/* エディタでこの関数を「実行」すると、台帳が読めるかどうかを確かめられます */
function testNippo() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const sh = nippoSheet_();
  Logger.log('日報台帳：' + (sh.getLastRow() - 1) + '件。今日（' + today + '）の現場日報：' + readDay_(today, null).length + '件');
}

/* 日報アプリの「日報台帳」を探す（見つけたら場所を覚えて、次からはすぐ開く） */
function nippoSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('NIPPO_SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) {} }
  const it = DriveApp.getFilesByName(NIPPO_SHEET);
  while (it.hasNext()) {
    const f = it.next();
    if (f.getMimeType() !== 'application/vnd.google-apps.spreadsheet' || f.isTrashed()) continue;
    const sh = SpreadsheetApp.openById(f.getId()).getSheets()[0];
    if (String(sh.getRange(1, 1).getValue()) !== 'recordId') continue;      // 日報アプリの台帳かどうか
    props.setProperty('NIPPO_SHEET_ID', f.getId());
    return sh;
  }
  throw new Error('日報台帳が見つかりません');
}

/* その日の現場日報を読む。name があれば、その人が入っているものだけ。
 * まず「作業日」と「削除」の列だけを読み、当てはまる行だけ中身を読む（台帳が大きくなっても遅くならないように）。 */
function readDay_(date, name) {
  const sh = nippoSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = last - 1;
  const days = sh.getRange(2, NIPPO_COL.date, n, 1).getDisplayValues();
  const dels = sh.getRange(2, NIPPO_COL.deleted, n, 1).getDisplayValues();
  const out = [];
  for (let i = 0; i < n; i++) {
    if (String(days[i][0]).trim() !== date || String(dels[i][0]).trim() === '削除') continue;
    const row = sh.getRange(i + 2, 1, 1, NIPPO_COL.data).getValues()[0];
    let rec = null;
    const cell = String(row[NIPPO_COL.data - 1] || '');
    try { rec = cell ? JSON.parse(cell) : null; } catch (e) { rec = null; }
    if (!rec) {
      try { rec = JSON.parse(DriveApp.getFileById(String(row[NIPPO_COL.dataId - 1])).getBlob().getDataAsString('UTF-8')); } catch (e) { rec = null; }
    }
    if (!rec || !Array.isArray(rec.workers)) continue;
    if (name && !rec.workers.some(function (w) { return w && String(w.name || '').trim() === name; })) continue;
    rec = slim_(rec);
    rec.recordId = String(row[NIPPO_COL.id - 1]);
    if (!rec.date) rec.date = date;
    out.push(rec);
  }
  return out;
}

/* サインの画像などの長い文字は返さない */
function slim_(v) {
  if (typeof v === 'string') return v.length > 1500 ? '' : v;
  if (Array.isArray(v)) return v.map(slim_);
  if (v && typeof v === 'object') { const o = {}; Object.keys(v).forEach(function (k) { o[k] = slim_(v[k]); }); return o; }
  return v;
}
