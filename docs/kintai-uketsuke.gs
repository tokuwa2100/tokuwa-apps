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
const VERSION = '1.1.0';

/* ── 現場日報（コア日報・常用日報）を、個人日報の下書きに使うための設定 ──
 * 日報アプリの受け口のアドレス。事務所用の合言葉はここには書きません。
 * 「プロジェクトの設定」→「スクリプト プロパティ」に、
 *     プロパティ：NIPPO_OFFICE_KEY　　値：事務所用の合言葉
 * を1つ追加してください（集計表の「ドライブ連携の設定」に入れているものと同じ）。 */
const NIPPO_URL = 'https://script.google.com/macros/s/AKfycby3MD8Lj1om5jopPcdJLktTCk4eZ_bI5aRrmbhj9qDv3Cn9GLVfFL9qLfvOvL5tCLT_1A/exec';

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
    if (req.action === 'ping') return out_({ ok: true, email: email, folder: rootFolder_().getUrl(), version: VERSION, nippo: !!nippoKey_() });
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
 * 日報の中身は、ここにもドライブにも保存しません（読んで返すだけ）。
 * 毎回最初から全部を読まなくてすむように、「どこまで読んだか」の目印（日ごとに1つの短い文字）だけを覚えます。 */
function nippoKey_() {
  return PropertiesService.getScriptProperties().getProperty('NIPPO_OFFICE_KEY') || '';
}

function siteDay_(req) {
  const date = String(req.date || ''), name = String(req.name || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name) return { ok: false, error: '日付か名前がありません' };
  const key = nippoKey_();
  if (!key) return { ok: false, error: 'not_configured' };
  return { ok: true, date: date, records: readDay_(key, date, name) };
}

/* 最初に1回、エディタでこの関数を「実行」すると、つながるかどうかを確かめられます（合言葉を登録したあとで） */
function testNippo() {
  const key = nippoKey_();
  if (!key) { Logger.log('スクリプト プロパティ NIPPO_OFFICE_KEY がまだ登録されていません'); return; }
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const n = readDay_(key, today, null).length;
  Logger.log('つながりました。今日（' + today + '）の現場日報：' + n + '件');
}

/* その日の現場日報を読む。name があれば、その人が入っているものだけ。
 * 目印：その日より前の日に覚えた「読み終わりの位置」から読み始める（その日の日報は、それより後に届くため）。 */
function readDay_(key, date, name) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  let marks = {};
  try { marks = JSON.parse(props.getProperty('NIPPO_MARKS') || '{}') || {}; } catch (e) { marks = {}; }
  const before = Object.keys(marks).filter(function (d) { return d < date; }).sort();
  let since = before.length ? marks[before[before.length - 1]] : '';
  const found = {};
  for (let guard = 0; guard < 80; guard++) {
    const r = UrlFetchApp.fetch(NIPPO_URL, {
      method: 'post', contentType: 'text/plain;charset=utf-8',
      payload: JSON.stringify({ action: 'list', key: key, since: since, limit: 200 }), muteHttpExceptions: true
    });
    let res = null;
    try { res = JSON.parse(r.getContentText()); } catch (e) { res = null; }
    if (!res || !res.ok) throw new Error('nippo:' + (res && res.error ? res.error : 'network'));
    (res.items || []).forEach(function (it) {
      if (!it) return;
      const id = String(it.id);
      delete found[id];                                   // 直された・消された日報は、新しいほうで置きかえる
      if (it.deleted || !it.record || it.record.date !== date) return;
      if (name && !(it.record.workers || []).some(function (w) { return w && String(w.name || '').trim() === name; })) return;
      const rec = slim_(it.record);
      rec.recordId = id;
      found[id] = rec;
    });
    since = res.next || since;
    if (!res.more) break;
  }
  /* 今日の目印がまだなければ、いまの読み終わりの位置を覚える（古い目印は120日で消す） */
  if (!marks[today] && since) {
    marks[today] = since;
    const limit = Utilities.formatDate(new Date(Date.now() - 120 * 86400000), 'Asia/Tokyo', 'yyyy-MM-dd');
    Object.keys(marks).forEach(function (d) { if (d < limit) delete marks[d]; });
    props.setProperty('NIPPO_MARKS', JSON.stringify(marks));
  }
  return Object.keys(found).map(function (k) { return found[k]; });
}

/* サインの画像などの長い文字は返さない */
function slim_(v) {
  if (typeof v === 'string') return v.length > 1500 ? '' : v;
  if (Array.isArray(v)) return v.map(slim_);
  if (v && typeof v === 'object') { const o = {}; Object.keys(v).forEach(function (k) { o[k] = slim_(v[k]); }); return o; }
  return v;
}
