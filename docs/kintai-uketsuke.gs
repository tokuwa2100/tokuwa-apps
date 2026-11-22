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
    if (req.action === 'ping') return out_({ ok: true, email: email, folder: rootFolder_().getUrl() });
    if (req.action === 'saveMonth') return out_(saveMonth_(req, email));
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
