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
const NIPPO_CACHE = '_現場日報の控え（さわらない）';

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
 * ボードで作業員が日報を書く日を選ぶと、その日・その人の現場日報だけを返します。
 * 毎回ぜんぶ読み直さないように、「勤怠」フォルダの中に月ごとの控えを置き、前回からの差分だけを足します。 */
function nippoKey_() {
  return PropertiesService.getScriptProperties().getProperty('NIPPO_OFFICE_KEY') || '';
}

function siteDay_(req) {
  const date = String(req.date || ''), name = String(req.name || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name) return { ok: false, error: '日付か名前がありません' };
  const key = nippoKey_();
  if (!key) return { ok: false, error: 'not_configured' };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let month;
  try {
    syncNippo_(key);
    month = readJson_(date.slice(0, 7) + '.json');
  } finally { lock.releaseLock(); }
  const records = Object.keys(month).map(function (k) { return month[k]; }).filter(function (r) {
    return r && r.date === date && (r.workers || []).some(function (w) { return w && String(w.name || '').trim() === name; });
  });
  return { ok: true, date: date, records: records };
}

/* 最初に1回、エディタでこの関数を「実行」すると、つながるかどうかを確かめられます（合言葉を登録したあとで） */
function testNippo() {
  const key = nippoKey_();
  if (!key) { Logger.log('スクリプト プロパティ NIPPO_OFFICE_KEY がまだ登録されていません'); return; }
  syncNippo_(key);
  const ids = readJson_('_ids.json');
  Logger.log('現場日報の控え：' + Object.keys(ids).length + '件');
}

function syncNippo_(key) {
  const props = PropertiesService.getScriptProperties();
  let since = props.getProperty('NIPPO_SINCE') || '';
  const ids = readJson_('_ids.json');          // 日報のID → 何月の控えに入っているか
  const months = {};                            // 今回さわった月の控え
  const monthOf = function (ym) { if (!months[ym]) months[ym] = readJson_(ym + '.json'); return months[ym]; };
  let changed = false;
  for (let guard = 0; guard < 60; guard++) {
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
      const old = ids[id];
      if (old) { delete monthOf(old)[id]; delete ids[id]; changed = true; }
      if (it.deleted || !it.record || !it.record.date) return;
      const rec = slim_(it.record);
      rec.recordId = id;
      const ym = String(rec.date).slice(0, 7);
      monthOf(ym)[id] = rec; ids[id] = ym; changed = true;
    });
    since = res.next || since;
    if (!res.more) break;
  }
  if (changed) {
    Object.keys(months).forEach(function (ym) { writeJson_(ym + '.json', months[ym]); });
    writeJson_('_ids.json', ids);
  }
  props.setProperty('NIPPO_SINCE', since);
}

/* サインの画像などの長い文字は控えに入れない */
function slim_(v) {
  if (typeof v === 'string') return v.length > 1500 ? '' : v;
  if (Array.isArray(v)) return v.map(slim_);
  if (v && typeof v === 'object') { const o = {}; Object.keys(v).forEach(function (k) { o[k] = slim_(v[k]); }); return o; }
  return v;
}

function cacheFolder_() { return sub_(rootFolder_(), NIPPO_CACHE); }
function readJson_(name) {
  const it = cacheFolder_().getFilesByName(name);
  if (!it.hasNext()) return {};
  try { return JSON.parse(it.next().getBlob().getDataAsString('UTF-8')) || {}; } catch (e) { return {}; }
}
function writeJson_(name, obj) {
  const f = cacheFolder_(), body = JSON.stringify(obj);
  const it = f.getFilesByName(name);
  if (it.hasNext()) it.next().setContent(body); else f.createFile(name, body, 'application/json');
}

