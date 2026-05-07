var SHEET_ID = "1wjKFnzeKQe6StoIHXcE08qC2bZgzgJWfywHk32o0BSw";
var SHEET_NAME = "SL富里のコピー";

var LOG_PASSWORD = "hokuso2026";

var PLAN_FOLDER_NAME = "FAS計画生産シート";
var PLAN_SHEET_NAME = "計画シート管理";
var PLAN_DELETE_LOG = "計画シート削除ログ";

// apo.ozzio.info アポインター名 → FAS プランナー名 マッピング（23名、石井（俊）は除外）
var APO_TO_FAS_NAME = {
  "並木": "並木 孝博",
  "中渡瀬": "中渡瀬 理恵",
  "佐藤（凌）": "佐藤 凌太",
  "外山": "外山 由佳",
  "奥": "奥 崇晃",
  "守田": "守田 ゆき",
  "岡山": "岡山 純一",
  "後藤": "後藤 航太",
  "植野": "植野 友介",
  "武田": "武田 翔人",
  "渋谷": "渋谷 航大",
  "渡会": "渡会 佳祐",
  "田中（乃）": "田中 乃凪",
  "田中（大）": "田中 大也",
  "石井（真）": "石井 真弓",
  "福田": "福田 花梨",
  "鈴木": "鈴木 啓人",
  "須田": "須田 歩",
  "飯田（敬）": "飯田 敬介",
  "飯田（未）": "飯田 未来",
  "高橋": "高橋 昌則",
  "髙木": "髙木 柚伽",
  "齋藤": "齋藤 和樹"
};

// 月初リセット対象（当月項目）。ここに無い項目（FAS/WIN/イベント/人材獲得）は累計として継続
var MONTHLY_RESET_ITEMS = [
  // アセット実現（当月）13項目
  "VC","内AO生","内専門生","内ビジネス","内DXエントリー","内DXレギュラー",
  "CEX","従業員オプション","WEB","GoogleWokSpace","GenSpark","光/SIM/モバイルルータ","我が社",
  // アポイント（当月）4項目
  "アポイントコール数","アポイントラウンド数","アポイント（確定）軒数","アポイント（確定）人数"
];

function doGet(e) {
  var callback = e.parameter.callback || "";
  var action = e.parameter.action || "";
  var json = "";

  if (action === "save") {
    json = JSON.stringify(saveData(e));
  } else if (action === "totals") {
    json = JSON.stringify(readTotals());
  } else if (action === "log") {
    json = JSON.stringify(readLog(e));
  } else if (action === "mylog") {
    json = JSON.stringify(readMyLog(e));
  } else if (action === "plan_list") {
    json = JSON.stringify(planList(e.parameter.name || ""));
  } else if (action === "plan_delete") {
    json = JSON.stringify(planDelete(e.parameter.name || "", e.parameter.fileId || ""));
  } else if (action === "conflicts") {
    json = JSON.stringify(getConflicts());
  } else if (action === "apo_status") {
    json = JSON.stringify(getApoStatus());
  } else if (action === "notices") {
    json = JSON.stringify(getNotices());
  } else {
    json = JSON.stringify(readData(e));
  }

  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.action === "plan_upload") {
    var result = planUpload(data.name, data.title, data.filename, data.contentType, data.data);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }
  if (data.action === "sync_apo") {
    var result = syncFromApo(data);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }
  var result = saveDataDirect(data.name, data.values);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function readData(e) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAME);
  var allData = ws.getDataRange().getValues();
  var targetName = e.parameter.name || "";
  var nameRow = allData[2];
  var itemCol = 2;
  var result = {};
  for (var c = 3; c < nameRow.length; c++) {
    var name = nameRow[c] ? nameRow[c].toString().trim() : "";
    if (!name) continue;
    if (targetName && name.replace(/\s/g, "") !== targetName.replace(/\s/g, "")) continue;
    var person = {};
    for (var r = 3; r < allData.length; r++) {
      var item = allData[r][itemCol] ? allData[r][itemCol].toString().trim() : "";
      if (item) {
        person[item] = allData[r][c] !== null && allData[r][c] !== "" ? allData[r][c] : 0;
      }
    }
    result[name] = person;
  }
  return result;
}

function readTotals() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAME);
  var allData = ws.getDataRange().getValues();
  var nameRow = allData[2]; // 3行目 = 名前行
  var itemCol = 2;          // C列 = 項目名

  // 項目名→行番号のマッピング
  var itemRows = {};
  for (var r = 3; r < allData.length; r++) {
    var item = allData[r][itemCol] ? allData[r][itemCol].toString().trim() : "";
    if (item) itemRows[item] = r;
  }

  var totalCust = 0, totalDX = 0, totalAO = 0;
  var planners = {};

  for (var c = 3; c < nameRow.length; c++) {
    var name = nameRow[c] ? nameRow[c].toString().trim() : "";
    if (!name) continue;
    var nameKey = name.replace(/\s/g, "");
    var cust = Number(allData[itemRows["担当軒数"]] ? allData[itemRows["担当軒数"]][c] : 0) || 0;
    var people = Number(allData[itemRows["担当人数"]] ? allData[itemRows["担当人数"]][c] : 0) || 0;
    var ippan = Number(allData[itemRows["内一般聴講生軒数"]] ? allData[itemRows["内一般聴講生軒数"]][c] : 0) || 0;
    var senmon = Number(allData[itemRows["内専門聴講生軒数"]] ? allData[itemRows["内専門聴講生軒数"]][c] : 0) || 0;
    var dx = Number(allData[itemRows["内DX会員軒数"]] ? allData[itemRows["内DX会員軒数"]][c] : 0) || 0;
    var ao = ippan + senmon + dx;

    var assetKeys = ["VC","内AO生","内専門生","内ビジネス","内DXエントリー","内DXレギュラー","CEX","従業員オプション","WEB","GoogleWokSpace","GenSpark","光/SIM/モバイルルータ","我が社"];
    var asset = 0;
    for (var ak = 0; ak < assetKeys.length; ak++) {
      asset += Number(allData[itemRows[assetKeys[ak]]] ? allData[itemRows[assetKeys[ak]]][c] : 0) || 0;
    }

    totalCust += cust;
    totalDX += dx;
    totalAO += ao;
    planners[nameKey] = { cust: cust, people: people, ao: ao, dx: dx, asset: asset };
  }

  return { totalCust: totalCust, totalDX: totalDX, totalAO: totalAO, planners: planners };
}

function saveData(e) {
  var name = e.parameter.name || "";
  var dataStr = e.parameter.data || "{}";
  var values = JSON.parse(dataStr);
  return saveDataDirect(name, values);
}

function saveDataDirect(name, values) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAME);
  var names = ws.getRange(3, 1, 1, ws.getLastColumn()).getValues()[0];
  var col = -1;
  for (var i = 0; i < names.length; i++) {
    if (names[i] && names[i].toString().replace(/\s/g, "") === name.replace(/\s/g, "")) {
      col = i + 1;
      break;
    }
  }
  if (col === -1) {
    return {status: "error", message: "name not found: " + name};
  }
  var items = ws.getRange(1, 3, ws.getLastRow(), 1).getValues();
  var updates = [];
  var oldValues = {};
  var keys = Object.keys(values);
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    var value = values[key];
    for (var r = 0; r < items.length; r++) {
      if (items[r][0] && items[r][0].toString().trim() === key.trim()) {
        var oldVal = ws.getRange(r + 1, col).getValue();
        oldValues[key] = oldVal !== null && oldVal !== "" ? oldVal : 0;
        ws.getRange(r + 1, col).setValue(Number(value) || 0);
        updates.push(key);
        break;
      }
    }
  }
  ws.getRange(2, 2).setValue("更新日:" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy年M月d日"));

  // 変更ログを記録
  writeLog(ss, name, values, updates, oldValues);

  return {status: "ok", updated: updates, name: name};
}

// 全拠点のプランナーをスプレッドシートに追加（GASエディタで1回だけ実行）
function addAllPlanners() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAME);
  var nameRow = ws.getRange(3, 1, 1, ws.getLastColumn()).getValues()[0];

  // 既存の名前を取得（スペース除去で比較用）
  var existing = {};
  for (var i = 0; i < nameRow.length; i++) {
    if (nameRow[i]) existing[nameRow[i].toString().replace(/\s/g, "")] = true;
  }

  // 拠点ごとのプランナー一覧
  var regions = [
    { name: "SL鎌ケ谷", members: ["片桐 啓介","入江 渓太","新田 祐二","柴 由香里","玉井 なおこ","関根 瞳","井戸 昭人","金 鍾進","濱田 耕一","工藤 李紀"] },
    { name: "成田デジタル館", members: ["武田 翔人"] },
    { name: "DCL千葉NT", members: ["齋藤 えま","森川 直人","小野寺 真唯","秋山 徹平","佐藤 諒","原重 彩花","烏蘭 其其格","髙梨 聖真"] },
    { name: "DCL旭", members: ["平山 敬人","早川 浩史","小林 敬史","遠藤 晴香","柴 茂雄"] },
    { name: "DCL四街道", members: ["富井 優太","藤本 圭","星野 凌我","橘木 悠人"] },
    { name: "DCL白井駅前", members: ["二村 和徳","渡辺 久斗","椎名 淳之"] },
    { name: "SL成田", members: ["諏訪 開飛","小林 篤司","山野井 尋紀","保屋松 綾太"] }
  ];

  var nextCol = ws.getLastColumn() + 1;
  var added = [];

  for (var r = 0; r < regions.length; r++) {
    var region = regions[r];
    for (var m = 0; m < region.members.length; m++) {
      var member = region.members[m];
      var key = member.replace(/\s/g, "");
      if (existing[key]) continue; // 既存ならスキップ

      // 2行目に拠点名、3行目に名前を書き込む
      ws.getRange(2, nextCol).setValue(region.name);
      ws.getRange(3, nextCol).setValue(member);
      added.push(region.name + ": " + member);
      nextCol++;
    }
  }

  // SL富里の既存メンバーにも拠点名を追加（2行目が空なら）
  for (var c = 4; c <= ws.getLastColumn(); c++) {
    var row2 = ws.getRange(2, c).getValue();
    var row3 = ws.getRange(3, c).getValue();
    if (row3 && !row2) {
      ws.getRange(2, c).setValue("SL富里インター");
    }
  }

  return "追加完了: " + added.length + "名\n" + added.join("\n");
}

function readMyLog(e) {
  var name = e.parameter.name || "";
  if (!name) return { status: "error", message: "name required" };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var logSheet = ss.getSheetByName("変更ログ");
  if (!logSheet) return { status: "ok", logs: [] };
  var data = logSheet.getDataRange().getValues();
  var displayData = logSheet.getDataRange().getDisplayValues();
  var logs = [];
  for (var r = 1; r < data.length; r++) {
    var logName = data[r][1] ? data[r][1].toString().replace(/\s/g, "") : "";
    if (logName !== name.replace(/\s/g, "")) continue;
    var dateStr = displayData[r][0] || "";
    logs.push({
      date: dateStr,
      item: data[r][2] ? data[r][2].toString() : "",
      oldVal: data[r][3] !== null && data[r][3] !== "" ? data[r][3] : 0,
      newVal: data[r][4] !== null && data[r][4] !== undefined && data[r][4] !== "" ? data[r][4] : (data[r][3] !== null && data[r][3] !== "" ? data[r][3] : 0)
    });
  }
  return { status: "ok", logs: logs };
}

function readLog(e) {
  var pw = e.parameter.pw || "";
  if (pw !== LOG_PASSWORD) {
    return { status: "error", message: "パスワードが違います" };
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var logSheet = ss.getSheetByName("変更ログ");
  if (!logSheet) {
    return { status: "ok", logs: [] };
  }
  var data = logSheet.getDataRange().getValues();
  var displayData = logSheet.getDataRange().getDisplayValues();
  var logs = [];
  for (var r = 1; r < data.length; r++) {
    var dateStr = displayData[r][0] || "";
    logs.push({
      date: dateStr,
      name: data[r][1] ? data[r][1].toString() : "",
      item: data[r][2] ? data[r][2].toString() : "",
      oldVal: data[r][3] !== null && data[r][3] !== "" ? data[r][3] : 0,
      newVal: data[r][4] !== null && data[r][4] !== undefined && data[r][4] !== "" ? data[r][4] : (data[r][3] !== null && data[r][3] !== "" ? data[r][3] : 0)
    });
  }
  return { status: "ok", logs: logs };
}

function getOrCreatePlanFolder() {
  var folders = DriveApp.getFoldersByName(PLAN_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PLAN_FOLDER_NAME);
}

function getOrCreatePlanSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(PLAN_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PLAN_SHEET_NAME);
    sh.appendRow(["日時", "プランナー名", "タイトル", "ファイル名", "ファイルID", "MIMEタイプ", "状態"]);
    sh.getRange(1, 1, 1, 7).setFontWeight("bold");
  }
  return sh;
}

function planUpload(name, title, filename, contentType, base64Data) {
  if (!name || !title || !filename || !base64Data) {
    return { status: "error", message: "必要な項目が不足しています" };
  }
  try {
    var folder = getOrCreatePlanFolder();
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType || "application/octet-stream", filename);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var sh = getOrCreatePlanSheet();
    var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    sh.appendRow([now, name, title, filename, file.getId(), contentType || "", "active"]);
    return {
      status: "ok",
      fileId: file.getId(),
      date: now,
      title: title,
      filename: filename,
      contentType: contentType || ""
    };
  } catch (err) {
    return { status: "error", message: String(err) };
  }
}

function planList(name) {
  try {
    var sh = getOrCreatePlanSheet();
    var data = sh.getDataRange().getValues();
    var displayData = sh.getDataRange().getDisplayValues();
    var list = [];
    var nameKey = (name || "").replace(/\s/g, "");
    for (var r = 1; r < data.length; r++) {
      var rowName = data[r][1] ? data[r][1].toString().replace(/\s/g, "") : "";
      if (nameKey && rowName !== nameKey) continue;
      var status = data[r][6] || "active";
      if (status === "deleted") continue;
      list.push({
        date: displayData[r][0] || "",
        name: data[r][1] ? data[r][1].toString() : "",
        title: data[r][2] ? data[r][2].toString() : "",
        filename: data[r][3] ? data[r][3].toString() : "",
        fileId: data[r][4] ? data[r][4].toString() : "",
        contentType: data[r][5] ? data[r][5].toString() : ""
      });
    }
    // 新しい順
    list.reverse();
    return { status: "ok", items: list };
  } catch (err) {
    return { status: "error", message: String(err), items: [] };
  }
}

function planDelete(name, fileId) {
  if (!name || !fileId) {
    return { status: "error", message: "name と fileId が必要です" };
  }
  try {
    var sh = getOrCreatePlanSheet();
    var data = sh.getDataRange().getValues();
    var nameKey = name.replace(/\s/g, "");
    for (var r = 1; r < data.length; r++) {
      if (data[r][4] && data[r][4].toString() === fileId) {
        var rowName = data[r][1] ? data[r][1].toString().replace(/\s/g, "") : "";
        if (rowName !== nameKey) {
          return { status: "error", message: "本人のみ削除できます" };
        }
        sh.getRange(r + 1, 7).setValue("deleted");
        var title = data[r][2] ? data[r][2].toString() : "";
        var filename = data[r][3] ? data[r][3].toString() : "";
        try {
          var file = DriveApp.getFileById(fileId);
          file.setTrashed(true);
        } catch (e) {}
        writePlanDeleteLog(name, title, filename, fileId);
        return { status: "ok" };
      }
    }
    return { status: "error", message: "該当ファイルが見つかりません" };
  } catch (err) {
    return { status: "error", message: String(err) };
  }
}

function writePlanDeleteLog(name, title, filename, fileId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(PLAN_DELETE_LOG);
  if (!sh) {
    sh = ss.insertSheet(PLAN_DELETE_LOG);
    sh.appendRow(["日時", "プランナー名", "タイトル", "ファイル名", "ファイルID"]);
    sh.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  sh.appendRow([now, name, title, filename, fileId]);
}

// 月初リセット: 現シートを指定名にコピー保存し、当月項目を0にリセット
// 引数 archiveNameParam: 指定すればそれを使う。省略時はデフォルト名を使用（手動実行用）
function archiveAndResetMonth(archiveNameParam) {
  var archiveName = archiveNameParam || "2026年4月"; // ★手動実行時のデフォルト
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var src = ss.getSheetByName(SHEET_NAME);
  if (!src) return { status: "error", message: "現シートが見つかりません: " + SHEET_NAME };
  if (ss.getSheetByName(archiveName)) {
    return { status: "error", message: archiveName + " は既に存在します。アーカイブ名を変えるか既存を削除してください" };
  }

  // 1) アーカイブ作成（現シートを丸ごとコピー）
  var archive = src.copyTo(ss);
  archive.setName(archiveName);
  ss.setActiveSheet(archive);
  ss.moveActiveSheet(ss.getNumSheets());

  // 2) 現シートの当月項目を0リセット
  var lastCol = src.getLastColumn();
  var lastRow = src.getLastRow();
  var allData = src.getRange(1, 1, lastRow, lastCol).getValues();
  var monthlyMap = {};
  for (var i = 0; i < MONTHLY_RESET_ITEMS.length; i++) monthlyMap[MONTHLY_RESET_ITEMS[i]] = true;

  var resetRows = [];
  for (var r = 3; r < allData.length; r++) {
    var item = allData[r][2] ? allData[r][2].toString().trim() : "";
    if (item && monthlyMap[item]) resetRows.push(r);
  }

  // D列（4列目）以降をまとめて0で上書き
  var width = lastCol - 3;
  if (width > 0) {
    var zeros = [];
    for (var c = 0; c < width; c++) zeros.push(0);
    for (var k = 0; k < resetRows.length; k++) {
      src.getRange(resetRows[k] + 1, 4, 1, width).setValues([zeros]);
    }
  }

  src.getRange(2, 2).setValue("更新日:" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy年M月d日") + "（" + archiveName + "アーカイブ後リセット済）");

  return {
    status: "ok",
    archived: archiveName,
    resetItemCount: resetRows.length,
    message: archiveName + " にアーカイブし、" + resetRows.length + "項目を0リセットしました"
  };
}

// アーカイブシートから当月項目を現シートに復元する（リセットの取り消し用）
// ★GASエディタから手動実行する。下の archiveName を復元元のシート名に書き換えてから実行
function restoreFromArchive() {
  var archiveName = "2026年4月"; // ★復元元のアーカイブシート名
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var src = ss.getSheetByName(archiveName);
  var dst = ss.getSheetByName(SHEET_NAME);
  if (!src) return { status: "error", message: archiveName + " シートが見つかりません" };
  if (!dst) return { status: "error", message: SHEET_NAME + " シートが見つかりません" };

  var monthlyMap = {};
  for (var i = 0; i < MONTHLY_RESET_ITEMS.length; i++) monthlyMap[MONTHLY_RESET_ITEMS[i]] = true;

  var lastCol = Math.min(src.getLastColumn(), dst.getLastColumn());
  var width = lastCol - 3;
  if (width <= 0) return { status: "error", message: "プランナー列がありません" };

  var srcItems = src.getRange(1, 3, src.getLastRow(), 1).getValues();
  var restoredItems = 0;
  for (var r = 3; r < srcItems.length; r++) {
    var item = srcItems[r][0] ? srcItems[r][0].toString().trim() : "";
    if (!item || !monthlyMap[item]) continue;
    var values = src.getRange(r + 1, 4, 1, width).getValues();
    dst.getRange(r + 1, 4, 1, width).setValues(values);
    restoredItems++;
  }

  dst.getRange(2, 2).setValue("更新日:" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy年M月d日") + "（" + archiveName + "から当月項目を復元）");

  return {
    status: "ok",
    restored: restoredItems,
    message: archiveName + " から " + restoredItems + " 項目を復元しました"
  };
}

// 変更ログから当月分の純粋な実績値を計算して現シートに反映
// 各プランナー×当月項目について、最初の月内ログを基準に当月分を逆算する
//   - 月内ログなし → 0
//   - 最初のログが X→0 (X>0): 自己リセット → 当月値 = 現在値
//   - 最初のログが 0→Y: 0スタート → 当月値 = 現在値
//   - 最初のログが X→Y (Y>X>0): 上乗せ → 当月値 = 現在値 - X
//   - 最初のログが X→Y (X>Y>0): 修正/減算 → 当月値 = 現在値 - Y
//   - 結果が負になる場合は 0 にクランプ
// 引数 monthPrefixParam: 指定すればそれを使う。省略時はデフォルトを使用（手動実行用）
function recomputeMonthFromLogs(monthPrefixParam) {
  var monthPrefix = monthPrefixParam || "2026/05/"; // ★手動実行時のデフォルト
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var src = ss.getSheetByName(SHEET_NAME);
  var logSheet = ss.getSheetByName("変更ログ");
  if (!src) return { status: "error", message: SHEET_NAME + " シートが見つかりません" };
  if (!logSheet) return { status: "error", message: "変更ログシートが見つかりません" };

  var monthlyMap = {};
  for (var i = 0; i < MONTHLY_RESET_ITEMS.length; i++) monthlyMap[MONTHLY_RESET_ITEMS[i]] = true;

  var lastRow = src.getLastRow();
  var lastCol = src.getLastColumn();
  var srcData = src.getRange(1, 1, lastRow, lastCol).getValues();

  var itemToRow = {};
  for (var r = 3; r < srcData.length; r++) {
    var it = srcData[r][2] ? srcData[r][2].toString().trim() : "";
    if (it) itemToRow[it] = r;
  }

  // 変更ログから「各プランナー×当月項目の最初のエントリ」を抽出
  var logRange = logSheet.getDataRange();
  var logVals = logRange.getValues();
  var logDisp = logRange.getDisplayValues();
  var firstLog = {};
  for (var lr = 1; lr < logVals.length; lr++) {
    var date = logDisp[lr][0] || "";
    if (date.indexOf(monthPrefix) !== 0) continue;
    var nm = logVals[lr][1] ? logVals[lr][1].toString().replace(/\s/g, "") : "";
    var it = logVals[lr][2] ? logVals[lr][2].toString().trim() : "";
    if (!nm || !it || !monthlyMap[it]) continue;
    var key = nm + "|" + it;
    if (firstLog[key]) continue;
    firstLog[key] = {
      oldVal: Number(logVals[lr][3]) || 0,
      newVal: Number(logVals[lr][4]) || 0
    };
  }

  // 当月項目の行ごとに全プランナー列を一括更新
  var width = lastCol - 3;
  var rowsUpdated = 0;
  for (var mi = 0; mi < MONTHLY_RESET_ITEMS.length; mi++) {
    var it = MONTHLY_RESET_ITEMS[mi];
    var row = itemToRow[it];
    if (row === undefined) continue;
    var newRow = [];
    for (var c = 3; c < lastCol; c++) {
      var nm = srcData[2][c] ? srcData[2][c].toString().replace(/\s/g, "") : "";
      if (!nm) { newRow.push(srcData[row][c]); continue; }
      var current = Number(srcData[row][c]) || 0;
      var key = nm + "|" + it;
      var log = firstLog[key];
      var nv;
      if (!log) nv = 0;
      else if (log.newVal === 0 || log.oldVal === 0) nv = current;
      else if (log.newVal > log.oldVal) nv = current - log.oldVal;
      else nv = current - log.newVal;
      if (nv < 0) nv = 0;
      newRow.push(nv);
    }
    src.getRange(row + 1, 4, 1, width).setValues([newRow]);
    rowsUpdated++;
  }

  src.getRange(2, 2).setValue("更新日:" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy年M月d日") + "（" + monthPrefix + " ログから当月分を再計算）");

  return {
    status: "ok",
    rowsUpdated: rowsUpdated,
    message: monthPrefix + " のログから " + rowsUpdated + " 行を再計算しました"
  };
}

// apo.ozzio.info からの月次データを受け取って FAS スプレッドシートに反映（差分同期方式）
// payload: { action:"sync_apo", month:"2026/05", data:[{apo_name, values:{...}}, ...], syncDate:"..." }
// values: アポイントコール数 / アポイント（確定）軒数 / アポイント（確定）人数
//
// 動作:
//   - 「apo同期状態」シートに前回apo値を保存
//   - 今回apo値 - 前回apo値 = delta を計算し、FASセルに加算（上書きせず）
//   - これによりプランナーの手入力分は保持される
//   - 初回（状態なし）はブートストラップ:
//       FAS=0 かつ apo>0 → FASに apo値を書き込み（初期投入）
//       FAS>0 → FASは触らず状態のみ記録（既に同期済みと仮定）
function syncFromApo(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAME);
  if (!ws) return { status: "error", message: SHEET_NAME + " シートが見つかりません" };

  var data = payload.data || [];
  if (!data.length) return { status: "error", message: "データが空です" };

  // 名前→列マップ
  var names = ws.getRange(3, 1, 1, ws.getLastColumn()).getValues()[0];
  var nameToCol = {};
  for (var i = 0; i < names.length; i++) {
    if (names[i]) nameToCol[names[i].toString().replace(/\s/g, "")] = i + 1;
  }
  // 項目→行マップ
  var items = ws.getRange(1, 3, ws.getLastRow(), 1).getValues();
  var itemToRow = {};
  for (var r = 0; r < items.length; r++) {
    if (items[r][0]) itemToRow[items[r][0].toString().trim()] = r + 1;
  }

  // apo同期状態シートを取得して前回値マップを作成
  var stateSheet = getOrCreateApoStateSheet(ss);
  var stateData = stateSheet.getDataRange().getValues();
  var stateMap = {}; // key = "nameKey|item" → {row, prevApo}
  for (var sr = 1; sr < stateData.length; sr++) {
    var snm = stateData[sr][0] ? stateData[sr][0].toString().replace(/\s/g, "") : "";
    var sit = stateData[sr][1] ? stateData[sr][1].toString().trim() : "";
    if (!snm || !sit) continue;
    stateMap[snm + "|" + sit] = { row: sr + 1, prevApo: Number(stateData[sr][2]) || 0 };
  }

  var processed = 0, skipped = 0, conflictsDetected = 0, bootstrapped = 0;
  var skippedNames = [];
  var nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  for (var d = 0; d < data.length; d++) {
    var entry = data[d];
    var apoName = entry.apo_name || "";
    var fasName = APO_TO_FAS_NAME[apoName];
    if (!fasName) {
      skipped++;
      skippedNames.push(apoName + "（マッピングなし）");
      continue;
    }
    var nameKey = fasName.replace(/\s/g, "");
    var col = nameToCol[nameKey];
    if (!col) {
      skipped++;
      skippedNames.push(apoName + "→" + fasName + "（FAS列なし）");
      continue;
    }

    var values = entry.values || {};
    var updatedKeys = [];
    var oldValues = {};
    var newValuesForLog = {};
    var keys = Object.keys(values);

    for (var k = 0; k < keys.length; k++) {
      var item = keys[k];
      var row = itemToRow[item];
      if (!row) continue;

      var newApoVal = Number(values[item]) || 0;
      var stateKey = nameKey + "|" + item;
      var existing = stateMap[stateKey];
      var currentFasVal = Number(ws.getRange(row, col).getValue()) || 0;

      if (!existing) {
        // ブートストラップ（初回のみ「上書き」モード）:
        // FAS値をapo現在値に揃え、状態シートに記録する。
        // これにより、上書き同期からの移行時に取り残された差分が解消される。
        // 以降は通常の差分加算モードで動く。
        if (newApoVal !== currentFasVal) {
          oldValues[item] = currentFasVal;
          ws.getRange(row, col).setValue(newApoVal);
          updatedKeys.push(item);
          newValuesForLog[item] = newApoVal;
        }
        stateSheet.appendRow([fasName, item, newApoVal, nowStr]);
        bootstrapped++;
        continue;
      }

      // 差分同期: delta = 今回apo - 前回apo
      var delta = newApoVal - existing.prevApo;
      if (delta !== 0) {
        var newFasVal = currentFasVal + delta;
        if (newFasVal < 0) newFasVal = 0;
        if (newFasVal !== currentFasVal) {
          oldValues[item] = currentFasVal;
          ws.getRange(row, col).setValue(newFasVal);
          updatedKeys.push(item);
          newValuesForLog[item] = newFasVal;

          // 競合検知
          if (detectConflict(ss, fasName, item, "apo")) conflictsDetected++;
        }
        // 状態を更新
        stateSheet.getRange(existing.row, 3).setValue(newApoVal);
        stateSheet.getRange(existing.row, 4).setValue(nowStr);
        existing.prevApo = newApoVal;
      }
    }

    if (updatedKeys.length > 0) {
      writeLog(ss, fasName, newValuesForLog, updatedKeys, oldValues, "apo");
      processed++;
    }
  }

  ws.getRange(2, 2).setValue("更新日:" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy年M月d日 HH:mm") + "（apo差分同期）");

  return {
    status: "ok",
    processed: processed,
    bootstrapped: bootstrapped,
    skipped: skipped,
    skippedNames: skippedNames,
    conflicts: conflictsDetected,
    syncDate: payload.syncDate || ""
  };
}

// apo同期状態シートを取得（なければ作成）
function getOrCreateApoStateSheet(ss) {
  var sh = ss.getSheetByName("apo同期状態");
  if (!sh) {
    sh = ss.insertSheet("apo同期状態");
    sh.appendRow(["プランナー名", "項目名", "前回apo値", "最終同期日時"]);
    sh.getRange(1, 1, 1, 4).setFontWeight("bold");
  }
  return sh;
}

// apo同期状態をクリア（月初リセット時に呼ぶ）
function clearApoState(ss) {
  var sh = ss.getSheetByName("apo同期状態");
  if (!sh) return 0;
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 4).clearContent();
    return lastRow - 1;
  }
  return 0;
}

// 競合検知: 同じ項目が今日「manual と apo」両方で更新されたか確認、あれば「競合ログ」シートに記録
// 戻り値: 競合を検出して記録したら true
function detectConflict(ss, name, item, currentSource) {
  var logSheet = ss.getSheetByName("変更ログ");
  if (!logSheet) return false;
  var data = logSheet.getDataRange().getValues();
  var disp = logSheet.getDataRange().getDisplayValues();
  var todayStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  var oppositeSource = (currentSource === "apo") ? "manual" : "apo";
  var nameKey = name.replace(/\s/g, "");

  for (var r = data.length - 1; r >= 1; r--) {
    var dateStr = disp[r][0] || "";
    if (dateStr.indexOf(todayStr) !== 0) continue; // 今日のログのみ対象
    var logName = data[r][1] ? data[r][1].toString().replace(/\s/g, "") : "";
    if (logName !== nameKey) continue;
    var logItem = data[r][2] ? data[r][2].toString() : "";
    if (logItem !== item) continue;
    var logSource = data[r][5] ? data[r][5].toString() : "manual";
    if (logSource === oppositeSource) {
      // 競合発生 → 競合ログに記録
      recordConflict(ss, todayStr, name, item, currentSource, oppositeSource);
      return true;
    }
  }
  return false;
}

function recordConflict(ss, dateStr, name, item, sourceA, sourceB) {
  var sh = ss.getSheetByName("競合ログ");
  if (!sh) {
    sh = ss.insertSheet("競合ログ");
    sh.appendRow(["検出日時", "対象日", "プランナー", "項目", "ソース1", "ソース2", "状態"]);
    sh.getRange(1, 1, 1, 7).setFontWeight("bold");
  }
  // 既に同じ (対象日, プランナー, 項目) で active な競合があればスキップ
  var data = sh.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (data[r][1] === dateStr && data[r][2] === name && data[r][3] === item && (data[r][6] || "active") === "active") {
      return;
    }
  }
  var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  sh.appendRow([now, dateStr, name, item, sourceA, sourceB, "active"]);
}

// 競合一覧取得（active のみ）
function getConflicts() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName("競合ログ");
  if (!sh) return { status: "ok", conflicts: [] };
  var data = sh.getDataRange().getValues();
  var conflicts = [];
  for (var r = 1; r < data.length; r++) {
    var status = data[r][6] || "active";
    if (status !== "active") continue;
    conflicts.push({
      detectedAt: data[r][0] ? data[r][0].toString() : "",
      date: data[r][1] ? data[r][1].toString() : "",
      name: data[r][2] ? data[r][2].toString() : "",
      item: data[r][3] ? data[r][3].toString() : "",
      sources: [(data[r][4] || "").toString(), (data[r][5] || "").toString()]
    });
  }
  return { status: "ok", conflicts: conflicts };
}

// 連絡事項取得（「連絡事項」シートから、表示列が"active"または空の行のみ、新しい順）
// シートが無ければ自動作成
function getNotices() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName("連絡事項");
  if (!sh) {
    sh = ss.insertSheet("連絡事項");
    sh.appendRow(["日付", "内容", "表示"]);
    sh.getRange(1, 1, 1, 3).setFontWeight("bold");
    sh.getRange("A:A").setNumberFormat("yyyy/MM/dd");
    // 初期データ
    var today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
    sh.appendRow([today, "連絡事項機能を追加しました", "active"]);
  }
  var data = sh.getDataRange().getDisplayValues();
  var notices = [];
  for (var r = 1; r < data.length; r++) {
    var hidden = (data[r][2] || "").toString().trim();
    if (hidden && hidden !== "active" && hidden !== "表示") continue;
    var date = data[r][0] ? data[r][0].toString() : "";
    var content = data[r][1] ? data[r][1].toString() : "";
    if (!content) continue;
    notices.push({ date: date, content: content });
  }
  // 新しい順（追加順の逆）
  notices.reverse();
  return { status: "ok", notices: notices };
}

// apo同期ステータス取得（最終同期時刻、登録ユーザー数等）
function getApoStatus() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName(SHEET_NAME);
  var statusCell = ws ? ws.getRange(2, 2).getValue() : "";
  var apoUsers = [];
  for (var apoName in APO_TO_FAS_NAME) apoUsers.push(APO_TO_FAS_NAME[apoName]);
  return {
    status: "ok",
    apoUserCount: apoUsers.length,
    apoUsers: apoUsers,
    lastUpdate: statusCell ? statusCell.toString() : ""
  };
}

// ─── apo.ozzio.info 外部API版（GAS直接呼出版、apo_sync.py/launchd不要）─────

// APIキーをScriptProperties（GAS暗号化ストア）に保存
// ★初回のみGASエディタから setApoApiKey 関数を実行する
//   または別途 setApoApiKeyManual から手動設定
function setApoApiKey(key) {
  if (!key) {
    // 引数省略時はここに直書きしたキーを使う（手動実行用）
    key = "PUT_YOUR_API_KEY_HERE";
  }
  if (key === "PUT_YOUR_API_KEY_HERE") {
    return { status: "error", message: "APIキーが設定されていません。setApoApiKeyの第1引数または関数内defaultに実キーを入れて実行してください" };
  }
  PropertiesService.getScriptProperties().setProperty("APO_API_KEY", key);
  return { status: "ok", message: "APIキーを保存しました（先頭8文字: " + key.substring(0, 8) + "...）" };
}

function getApoApiKey() {
  return PropertiesService.getScriptProperties().getProperty("APO_API_KEY");
}

// apo外部APIの1人分レコードをFASのアポイント値にマッピング（旧: transform_apo_to_fas相当）
function transformApoRowToFasValues(row) {
  var outCalls = Number(row.total_outbound_calls) || 0;
  var inCalls = Number(row.total_inbound_calls) || 0;
  var outPromises = Number(row.total_outbound_visit_promises) || 0;
  var inPromises = Number(row.total_inbound_visit_promises) || 0;
  var outVisitors = Number(row.total_outbound_expected_visitors) || 0;
  var inVisitors = Number(row.total_inbound_expected_visitors) || 0;
  return {
    "アポイントコール数": outCalls + inCalls,
    "アポイント（確定）軒数": outPromises + inPromises,
    "アポイント（確定）人数": outVisitors + inVisitors
  };
}

// apo外部APIから当月分を取得してFASに同期する（時間トリガーから呼ばれる）
function fetchApoMonthlyAndSync() {
  var apiKey = getApoApiKey();
  if (!apiKey) {
    return { status: "error", message: "APIキー未設定。setApoApiKey('...')をGASエディタから1回実行してください" };
  }

  var now = new Date();
  // JST時刻計算
  var jst = new Date(now.getTime() + 9 * 3600 * 1000);
  var year = jst.getUTCFullYear();
  var month = jst.getUTCMonth() + 1;

  var url = "https://apo.ozzio.info/api/external/monthly?year=" + year + "&month=" + month;
  var response;
  try {
    response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        "X-API-Key": apiKey,
        "User-Agent": "FAS-Sync/1.0 (GAS)"
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
  } catch (e) {
    return _logApoFetchResult({ status: "error", message: "fetchエラー: " + e });
  }

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) {
    return _logApoFetchResult({ status: "error", message: "HTTP " + code + ": " + body.substring(0, 200) });
  }

  var json;
  try { json = JSON.parse(body); } catch (e) {
    return _logApoFetchResult({ status: "error", message: "JSON parseエラー: " + body.substring(0, 200) });
  }
  if (!json.success || !Array.isArray(json.data)) {
    return _logApoFetchResult({ status: "error", message: "予期しないレスポンス形式: " + body.substring(0, 200) });
  }

  // syncFromApo の payload 形式に変換
  var syncData = [];
  for (var i = 0; i < json.data.length; i++) {
    var row = json.data[i];
    var planner = row.planner_name || row.appointer_name || ""; // 旧API互換も一応
    if (!planner) continue;
    syncData.push({ apo_name: planner, values: transformApoRowToFasValues(row) });
  }

  var payload = {
    action: "sync_apo",
    month: year + "/" + ("0" + month).slice(-2),
    data: syncData,
    syncDate: Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss")
  };

  var result = syncFromApo(payload);
  return _logApoFetchResult(result);
}

function _logApoFetchResult(result) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName("自動実行ログ");
  if (!sh) {
    sh = ss.insertSheet("自動実行ログ");
    sh.appendRow(["実行日時", "対象月", "結果", "詳細"]);
    sh.getRange(1, 1, 1, 4).setFontWeight("bold");
  }
  var nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  var summary;
  if (result && result.status === "ok") {
    summary = "反映" + (result.processed || 0) + "/ブートストラップ" + (result.bootstrapped || 0) + "/スキップ" + (result.skipped || 0) + "/競合" + (result.conflicts || 0);
  } else {
    summary = (result && result.message) || "不明";
  }
  sh.appendRow([nowStr, "apo同期", (result && result.status) || "?", summary]);
  return result;
}

// 強制フル同期: apo現在値でFASを完全上書きし、状態シートをリセット。1回限りの修復用
// ★この後の自動sync（差分加算）は通常通り動く
function forceFullResyncFromApo() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  // 状態シート削除（次回 sync で空状態から再作成 → bootstrap が全項目に走る）
  var sh = ss.getSheetByName("apo同期状態");
  if (sh) ss.deleteSheet(sh);
  // 即時同期実行
  return fetchApoMonthlyAndSync();
}

// 自動同期トリガー設定: 1時間ごとに fetchApoMonthlyAndSync を実行
// ★GASエディタから1回だけ実行
function setupAutoApoSync() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "fetchApoMonthlyAndSync") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  ScriptApp.newTrigger("fetchApoMonthlyAndSync")
    .timeBased()
    .everyHours(1)
    .create();
  return {
    status: "ok",
    removed: removed,
    message: "1時間ごとに fetchApoMonthlyAndSync を自動実行するトリガーを設定しました"
  };
}

// 自動同期トリガーを削除
function removeAutoApoSync() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "fetchApoMonthlyAndSync") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  return { status: "ok", removed: removed, message: removed + " 個のトリガーを削除しました" };
}

// 自動月次リセット: 毎月1日 03:00 JST に時間トリガーから呼ばれる
// 前月分を再計算→アーカイブ→現シートを0リセット の順で実行
function monthlyResetTrigger() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var now = new Date();
  // 前月（実行日の1ヶ月前の月初）
  var prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var year = prev.getFullYear();
  var month = prev.getMonth() + 1;
  var monthPrefix = year + "/" + ("0" + month).slice(-2) + "/";
  var archiveName = year + "年" + month + "月";

  var execLog = ss.getSheetByName("自動実行ログ");
  if (!execLog) {
    execLog = ss.insertSheet("自動実行ログ");
    execLog.appendRow(["実行日時", "対象月", "結果", "詳細"]);
    execLog.getRange(1, 1, 1, 4).setFontWeight("bold");
  }
  var nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  try {
    // 1. 前月分を純粋な実績値に再計算（上書き派の累計を当月分のみに変換）
    var rec = recomputeMonthFromLogs(monthPrefix);
    if (rec.status !== "ok") throw new Error("recompute失敗: " + rec.message);

    // 2. アーカイブ＋リセット
    var arc = archiveAndResetMonth(archiveName);
    if (arc.status !== "ok") throw new Error("archive失敗: " + arc.message);

    // 3. apo同期状態をクリア（月境界で前回値をリセットしないと差分計算が破綻）
    var clearedRows = clearApoState(ss);

    execLog.appendRow([nowStr, archiveName, "成功", "再計算" + rec.rowsUpdated + "行 / " + arc.message + " / apo状態" + clearedRows + "行クリア"]);
    return { status: "ok", archiveName: archiveName, monthPrefix: monthPrefix };
  } catch (err) {
    execLog.appendRow([nowStr, archiveName, "失敗", String(err)]);
    return { status: "error", message: String(err) };
  }
}

// 自動トリガーを設定: 毎月1日 03:00 JST に monthlyResetTrigger を実行
// ★GASエディタから1回だけ手動実行すれば、以降は自動で走る
function setupAutoMonthlyReset() {
  // 既存の同名トリガーを削除して重複を防ぐ
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "monthlyResetTrigger") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  // 毎月1日 03:00 JST にトリガー設定
  ScriptApp.newTrigger("monthlyResetTrigger")
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .inTimezone("Asia/Tokyo")
    .create();
  return {
    status: "ok",
    removed: removed,
    message: "毎月1日 03:00 (JST) に monthlyResetTrigger を自動実行するトリガーを設定しました"
  };
}

// 自動トリガーを削除（自動運用を止める時に使う）
function removeAutoMonthlyReset() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "monthlyResetTrigger") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  return { status: "ok", removed: removed, message: removed + " 個のトリガーを削除しました" };
}

function writeLog(ss, name, values, updatedKeys, oldValues, source) {
  var logSheet = ss.getSheetByName("変更ログ");
  if (!logSheet) {
    logSheet = ss.insertSheet("変更ログ");
    logSheet.appendRow(["日時", "名前", "変更項目", "変更前", "変更後", "ソース"]);
    logSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  // ヘッダー整備（4列→5列→6列の段階アップグレード）
  var headerVals = logSheet.getRange(1, 1, 1, 6).getValues()[0];
  if (!headerVals[3] || headerVals[3] === "変更後の値") {
    logSheet.getRange(1, 4).setValue("変更前");
    logSheet.getRange(1, 5).setValue("変更後");
  }
  if (!headerVals[5]) {
    logSheet.getRange(1, 6).setValue("ソース");
  }
  logSheet.getRange(1, 1, 1, 6).setFontWeight("bold");

  var src = source || "manual";
  var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  for (var i = 0; i < updatedKeys.length; i++) {
    var key = updatedKeys[i];
    var oldVal = oldValues && oldValues[key] !== undefined ? oldValues[key] : "";
    logSheet.appendRow([now, name, key, oldVal, values[key], src]);
  }
}
