/* =========================================================
   徳和工業　日報の受け口（Googleドライブ保存）
   ---------------------------------------------------------
   ・日報アプリ（コア・常用）から届いた日報を、作業日ごとのフォルダに保存する
       日報（アプリ）/2026/2026-10/2026-10-30/ 日付_取引先_現場_作業員_コア_時刻.png / .json
   ・届いた日報の一覧は「日報台帳」（スプレッドシート）に1行ずつ記録する
   ・集計表で直した内容は、ドライブのデータ（.json）にも反映する
       （現場で署名された画像（.png）は原本としてそのまま残す。最初の修正時に「_原本.json」も残す）

   合言葉（このファイルには書かない）
     「プロジェクトの設定」→「スクリプト プロパティ」に次の2つを登録する
       WORKER_KEY … 作業員のスマホ（日報アプリ）に入れる合言葉。送信だけできる
       OFFICE_KEY … 事務所の集計表に入れる合言葉。一覧の取得・修正の反映ができる
   ========================================================= */

const VERSION    = "1.0.0";
const ROOT_NAME  = "日報（アプリ）";
const SHEET_NAME = "日報台帳";
const HEADERS = ["recordId", "種別", "作業日", "取引先", "現場", "作業員",
                 "受信日時", "更新日時", "更新者", "削除", "画像ID", "データID", "原本ID", "データ"];
const COL = {};
HEADERS.forEach(function (h, i) { COL[h] = i + 1; });
const CELL_MAX = 45000;   // スプレッドシートの1セルに入る文字数の目安（超えたら .json から読む）

/* ---------------- 入口 ---------------- */
function doGet() {
  return out_({ ok: true, app: "tokuwa-nippo-drive", version: VERSION });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ ok: false, error: "bad_request" }); }

  const props = PropertiesService.getScriptProperties();
  const workerKey = props.getProperty("WORKER_KEY");
  const officeKey = props.getProperty("OFFICE_KEY");
  if (!workerKey || !officeKey) return out_({ ok: false, error: "not_configured" });

  const isOffice = !!req.key && req.key === officeKey;
  const isWorker = isOffice || (!!req.key && req.key === workerKey);
  if (!isWorker) {
    Utilities.sleep(1500);   // 合言葉の総当たりを遅らせる
    return out_({ ok: false, error: "bad_key" });
  }

  try {
    switch (req.action) {
      case "ping":   return out_({ ok: true, role: isOffice ? "office" : "worker", version: VERSION });
      case "submit": return out_(submit_(req));
      case "list":   return isOffice ? out_(list_(req))   : out_({ ok: false, error: "office_only" });
      case "update": return isOffice ? out_(update_(req)) : out_({ ok: false, error: "office_only" });
      default:       return out_({ ok: false, error: "unknown_action" });
    }
  } catch (err) {
    console.error(err);
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 受信（日報アプリから） ---------------- */
// req: { record, image(base64 PNG・任意) }
function submit_(req) {
  const rec = req.record;
  if (!rec || typeof rec !== "object") return { ok: false, error: "no_record" };
  const id = recId_(rec);
  if (!id) return { ok: false, error: "bad_id" };

  return withLock_(function () {
    const sh = sheet_();
    const row = findRow_(sh, id);
    if (row) {
      // 同じ日報がもう届いている（再送）。画像だけ無ければ足す。内容（修正済みかもしれない）は上書きしない
      if (req.image && !sh.getRange(row, COL["画像ID"]).getValue()) {
        const folder = DriveApp.getFileById(String(sh.getRange(row, COL["データID"]).getValue())).getParents().next();
        const png = savePng_(folder, baseName_(rec), req.image);
        sh.getRange(row, COL["画像ID"]).setValue(png.getId());
        return { ok: true, status: "image_added" };
      }
      return { ok: true, status: "exists" };
    }

    const folder = dayFolder_(rec.date);
    const base = baseName_(rec);
    const png = req.image ? savePng_(folder, base, req.image) : null;
    const json = folder.createFile(base + ".json", JSON.stringify(rec, null, 1), "application/json");
    const now = new Date().toISOString();
    appendRow_(sh, id, rec, {
      "受信日時": now, "更新日時": now, "更新者": "（現場）", "削除": "",
      "画像ID": png ? png.getId() : "", "データID": json.getId(), "原本ID": ""
    });
    return { ok: true, status: "saved" };
  });
}

/* ---------------- 一覧（集計表へ） ---------------- */
// req: { since: 前回の更新日時（ISO・空なら全部）, limit }
function list_(req) {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, items: [], more: false, next: req.since || "" };
  const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const since = String(req.since || "");
  const limit = Math.min(Math.max(parseInt(req.limit, 10) || 200, 1), 500);

  let rows = vals.filter(function (v) { return String(v[COL["更新日時"] - 1]) > since; });
  rows.sort(function (a, b) {
    const x = String(a[COL["更新日時"] - 1]), y = String(b[COL["更新日時"] - 1]);
    return x < y ? -1 : x > y ? 1 : 0;
  });
  let more = false;
  if (rows.length > limit) {
    // 同じ更新日時の途中で切らない
    let n = limit;
    const edge = String(rows[n - 1][COL["更新日時"] - 1]);
    while (n < rows.length && String(rows[n][COL["更新日時"] - 1]) === edge) n++;
    more = n < rows.length;
    rows = rows.slice(0, n);
  }

  const items = rows.map(function (v) {
    let rec = null;
    const cell = String(v[COL["データ"] - 1] || "");
    try { rec = cell ? JSON.parse(cell) : null; } catch (e) { rec = null; }
    if (!rec) {
      try { rec = stripSig_(JSON.parse(DriveApp.getFileById(String(v[COL["データID"] - 1])).getBlob().getDataAsString())); }
      catch (e) { rec = null; }
    }
    return {
      id: String(v[0]),
      updatedAt: String(v[COL["更新日時"] - 1]),
      receivedAt: String(v[COL["受信日時"] - 1]),
      deleted: v[COL["削除"] - 1] === "削除",
      imageId: String(v[COL["画像ID"] - 1] || ""),
      record: rec
    };
  }).filter(function (it) { return it.record; });

  return { ok: true, items: items, more: more, next: items.length ? items[items.length - 1].updatedAt : since };
}

/* ---------------- 修正の反映（集計表から） ---------------- */
// req: { recordId, record(修正後), deleted, by, reason, image(base64・任意：画像で取り込んだ日報) }
function update_(req) {
  const rec = req.record;
  const id = String(req.recordId || (rec && recId_(rec)) || "");
  if (!rec || !id || !/^[\w\-~.]{1,120}$/.test(id)) return { ok: false, error: "bad_request" };
  rec.id = id;
  if (!rec.recordId) rec.recordId = id;

  return withLock_(function () {
    const sh = sheet_();
    const now = new Date().toISOString();
    const by = String(req.by || "（事務所）").slice(0, 40);
    const del = !!req.deleted;
    let row = findRow_(sh, id);

    // ドライブにまだ無い日報（LINEで届いて集計表に入れたもの等）は、ここで新しく作る
    if (!row) {
      const folder = dayFolder_(rec.date);
      const base = (del ? "【削除】" : "") + baseName_(rec);
      const png = req.image ? savePng_(folder, base, req.image) : null;
      const json = folder.createFile(base + ".json", JSON.stringify(withMeta_(rec, req, now, by), null, 1), "application/json");
      appendRow_(sh, id, rec, {
        "受信日時": now, "更新日時": now, "更新者": by, "削除": del ? "削除" : "",
        "画像ID": png ? png.getId() : "", "データID": json.getId(), "原本ID": ""
      });
      return { ok: true, status: "created" };
    }

    const jsonFile = DriveApp.getFileById(String(sh.getRange(row, COL["データID"]).getValue()));
    let old = {};
    try { old = JSON.parse(jsonFile.getBlob().getDataAsString()) || {}; } catch (e) { old = {}; }

    // 最初の修正のときだけ、現場から届いた内容を「_原本.json」として残す
    let origId = String(sh.getRange(row, COL["原本ID"]).getValue() || "");
    if (!origId && !old._updatedAt) {
      const orig = jsonFile.makeCopy(jsonFile.getName().replace(/\.json$/, "") + "_原本.json");
      origId = orig.getId();
    }

    // 署名の画像データは集計表に無いので、元のファイルから引き継ぐ
    if (old.signaturePng && !rec.signaturePng) rec.signaturePng = old.signaturePng;
    jsonFile.setContent(JSON.stringify(withMeta_(rec, req, now, by), null, 1));

    // 作業日が変わったらフォルダを移す。名前（取引先・現場など）が変わったらファイル名も直す
    const folder = dayFolder_(rec.date);
    const base = (del ? "【削除】" : "") + baseName_(rec);
    const files = [[jsonFile, ".json"]];
    const pngId = String(sh.getRange(row, COL["画像ID"]).getValue() || "");
    if (pngId) files.push([DriveApp.getFileById(pngId), ".png"]);
    else if (req.image) {
      const png = savePng_(folder, base, req.image);
      sh.getRange(row, COL["画像ID"]).setValue(png.getId());
    }
    if (origId) files.push([DriveApp.getFileById(origId), "_原本.json"]);
    files.forEach(function (f) {
      const file = f[0];
      if (!isIn_(file, folder)) file.moveTo(folder);
      if (file.getName() !== base + f[1]) file.setName(base + f[1]);
    });

    writeRow_(sh, row, id, rec, { "更新日時": now, "更新者": by, "削除": del ? "削除" : "", "原本ID": origId });
    return { ok: true, status: "updated" };
  });
}

function withMeta_(rec, req, now, by) {
  const o = JSON.parse(JSON.stringify(rec));
  o._updatedAt = now;
  o._updatedBy = by;
  o._deleted = !!req.deleted;
  if (req.edits) o._edits = req.edits;
  return o;
}

/* ---------------- 台帳（スプレッドシート） ---------------- */
function sheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SHEET_ID");
  if (id) {
    try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) { id = null; }
  }
  const ss = SpreadsheetApp.create(SHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(root_());
  const sh = ss.getSheets()[0];
  sh.setName("台帳");
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold").setBackground("#e8f0fe");
  sh.getRange(1, 1, 1000, HEADERS.length).setNumberFormat("@");   // 日付などを勝手に変換させない
  sh.setFrozenRows(1);
  sh.setColumnWidth(COL["データ"], 120);
  props.setProperty("SHEET_ID", ss.getId());
  return sh;
}

function findRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function rowValues_(id, rec, extra) {
  const data = JSON.stringify(stripSig_(rec));
  const v = {
    "recordId": id,
    "種別": kindLabel_(rec),
    "作業日": String(rec.date || ""),
    "取引先": clientOf_(rec),
    "現場": String(rec.site || ""),
    "作業員": workerNames_(rec).join("・"),
    "データ": data.length <= CELL_MAX ? data : ""
  };
  Object.keys(extra).forEach(function (k) { v[k] = extra[k]; });
  return v;
}

function appendRow_(sh, id, rec, extra) {
  const v = rowValues_(id, rec, extra);
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, HEADERS.length).setNumberFormat("@")
    .setValues([HEADERS.map(function (h) { return v[h] == null ? "" : String(v[h]); })]);
}

function writeRow_(sh, row, id, rec, extra) {
  const v = rowValues_(id, rec, extra);
  const cur = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  const next = HEADERS.map(function (h, i) { return v[h] == null ? cur[i] : String(v[h]); });
  sh.getRange(row, 1, 1, HEADERS.length).setValues([next]);
}

/* ---------------- フォルダ・ファイル ---------------- */
function root_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("ROOT_ID");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getRootFolder().getFoldersByName(ROOT_NAME);
  const f = it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(ROOT_NAME);
  props.setProperty("ROOT_ID", f.getId());
  return f;
}

function child_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// 作業日ごと：日報（アプリ）/2026/2026-10/2026-10-30
function dayFolder_(date) {
  const d = String(date || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return child_(root_(), "日付なし");
  return child_(child_(child_(root_(), d.slice(0, 4)), d.slice(0, 7)), d);
}

function savePng_(folder, base, b64) {
  const s = String(b64).replace(/^data:image\/\w+;base64,/, "");
  return folder.createFile(Utilities.newBlob(Utilities.base64Decode(s), "image/png", base + ".png"));
}

function isIn_(file, folder) {
  const it = file.getParents();
  while (it.hasNext()) if (it.next().getId() === folder.getId()) return true;
  return false;
}

/* ---------------- 日報の中身から取り出す ---------------- */
function recId_(rec) {
  const id = String(rec.recordId || rec.id || "");
  return /^[\w\-~.]{1,120}$/.test(id) ? id : "";
}
function isJoyou_(rec) { return rec.type === "joyou" || rec.app === "joyou-nippo"; }
function kindLabel_(rec) { return isJoyou_(rec) ? "常用" : "コア"; }
function clientOf_(rec) { return String(rec.client || rec.customer || ""); }
function workerNames_(rec) {
  return (rec.workers || []).map(function (w) { return w && w.name ? String(w.name).trim() : ""; }).filter(String);
}
function stripSig_(rec) {
  const o = JSON.parse(JSON.stringify(rec));
  delete o.signaturePng; delete o.sent; delete o.sentAt; delete o.driveSent; delete o.driveSentAt; delete o.driveSkip;
  return o;
}

// 日報アプリのファイル名と同じ形：日付_取引先_現場_作業員_コア_時刻
function baseName_(rec) {
  const clean = function (s, max) {
    const t = String(s == null ? "" : s).replace(/[\\\/:*?"<>|\r\n\t]/g, "").replace(/\s+/g, "").slice(0, max);
    return t || "なし";
  };
  const d = rec.date || "nodate";
  const t = (String(rec.createdAt || "").slice(11, 16).replace(":", "")) || "0000";
  const names = workerNames_(rec);
  let worker = names.length ? clean(names[0], 8) : "なし";
  if (names.length > 1) worker += "他" + (names.length - 1) + "名";
  return d + "_" + clean(clientOf_(rec), 12) + "_" + clean(rec.site, 16) + "_" + worker + "_" + kindLabel_(rec) + "_" + t;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- 初回の確認用（エディタから実行） ---------------- */
// 実行すると「日報（アプリ）」フォルダと「日報台帳」ができ、ドライブの使用を許可する画面が出る
function setupCheck() {
  const props = PropertiesService.getScriptProperties();
  console.log("フォルダ: " + root_().getUrl());
  console.log("台帳: " + sheet_().getParent().getUrl());
  console.log("WORKER_KEY: " + (props.getProperty("WORKER_KEY") ? "登録済み" : "★未登録"));
  console.log("OFFICE_KEY: " + (props.getProperty("OFFICE_KEY") ? "登録済み" : "★未登録"));
}
