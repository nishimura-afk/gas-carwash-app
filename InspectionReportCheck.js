/**
 * =====================================================
 * 洗車機点検報告書 自動チェックシステム
 * =====================================================
 *
 * 【概要】
 * Googleドライブの指定フォルダに保存されたアピカの点検報告書PDFを
 * 自動で読み取り、管理アプリの累計台数と比較してチェックする。
 *
 * 【処理の流れ】
 * 1. 指定フォルダ内の未処理PDFを検出
 * 2. PDFをGoogleドキュメントに変換してテキスト抽出（OCR）
 * 3. テキストから店舗名・累計台数を読み取り
 * 4. 管理アプリのデータ（ステータス集計）と比較（予測値との差分チェック）
 * 5. 結果をメールで通知 & PDFを自動リネーム+処理済みフォルダに移動
 *
 * 【セットアップ】
 * 1. GASエディタで「拡張機能」→「Google のサービス」→ Drive API を ON にし、
 *    Google Cloud Console で Drive API を有効化してください。
 * 2. setupFoldersInspectionReport() を一度実行して、フォルダを自動作成
 * 3. setupDailyTriggerInspectionReport() を実行して日次トリガーを設定
 *
 * 【運用】
 * フォルダ構成: 21_アピカ点検報告書 / アピカ点検報告書_受信 にPDFを入れると、
 * processInspectionReports() 実行時に処理され、処理済みは同親直下の アピカ点検報告書_処理済み へ移動します。
 */

// ============================================================
// 設定（点検報告書チェック専用）
// ============================================================
var INSPECTION_REPORT_CONFIG = {
  // Googleドライブのフォルダ構成（親フォルダの直下に受信・処理済み・一時変換）
  FOLDER_PARENT: "21_アピカ点検報告書",
  FOLDER_INBOX: "アピカ点検報告書_受信",
  FOLDER_DONE: "アピカ点検報告書_処理済み",
  // 一時変換: PDF→Google Doc に変換する際、Drive API が一時的に Doc を作成するための置き場。テキスト取得後に即削除する。
  TEMP_FOLDER: "アピカ点検_一時変換",

  // フォルダIDを指定すると名前検索を使わず確実にそのフォルダを参照する（共有ドライブ・同名フォルダ対策）
  FOLDER_PARENT_ID: "1OHnwdd_QiYRr1joibkNuCFqY-GsJ-zeN",
  FOLDER_INBOX_ID: "180DTMLrqSwElBwOuwqRElzCLNZjZHe9G",
  FOLDER_DONE_ID: "19AGnh4mDUvDKRt5QjWTpcbhkTd7zr9qn",

  // 通知先メールアドレス
  NOTIFY_EMAIL: "nishimura@selfix.jp",

  // 送信元メールアドレス（Gmail の「メールアドレスの追加」で登録済みであること）
  FROM_EMAIL: "nishimura@nishimurawork.com",

  // 累計台数のズレ許容範囲（月平均の何倍までOKか）1＝月平均以内
  THRESHOLD_MONTHS: 1
};

// ============================================================
// 初期セットアップ
// ============================================================

/**
 * 点検報告書用フォルダを自動作成する（最初に1回だけ実行）
 * 構成: 21_アピカ点検報告書 / アピカ点検報告書_受信, 処理済み, 一時変換
 */
function setupFoldersInspectionReport() {
  var root = DriveApp.getRootFolder();
  var parentName = INSPECTION_REPORT_CONFIG.FOLDER_PARENT;

  var parentIt = root.getFoldersByName(parentName);
  var parent;
  if (parentIt.hasNext()) {
    parent = parentIt.next();
    Logger.log("既存フォルダ（親）: " + parentName);
  } else {
    parent = root.createFolder(parentName);
    Logger.log("作成しました（親）: " + parentName + " (ID: " + parent.getId() + ")");
  }

  var subfolders = [
    INSPECTION_REPORT_CONFIG.FOLDER_INBOX,
    INSPECTION_REPORT_CONFIG.FOLDER_DONE,
    INSPECTION_REPORT_CONFIG.TEMP_FOLDER
  ];

  subfolders.forEach(function(name) {
    var it = parent.getFoldersByName(name);
    if (it.hasNext()) {
      Logger.log("既存: " + parentName + " / " + name);
    } else {
      parent.createFolder(name);
      Logger.log("作成しました: " + parentName + " / " + name);
    }
  });

  Logger.log("\n=== セットアップ完了 ===");
  Logger.log("「" + parentName + " / " + INSPECTION_REPORT_CONFIG.FOLDER_INBOX + "」にPDFを入れてください。");
}

/**
 * 点検報告書チェック用の日次トリガーを設定する（最初に1回だけ実行）
 * 毎日午前9時に自動実行。受信フォルダが常に処理され、空であることが多くなる。
 */
function setupDailyTriggerInspectionReport() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "processInspectionReports") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("processInspectionReports")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log("点検報告書チェックの日次トリガーを設定しました（毎日 9:00-10:00）");
}

// ============================================================
// メイン処理
// ============================================================

/**
 * メイン関数：フォルダ内のPDFを処理する
 * ※ 親・受信・処理済みを1回だけ取得し、同じフォルダ参照で移動する（同名フォルダが複数あるとずれないようにする）
 */
function processInspectionReports() {
  Logger.log("=== 点検報告書チェック開始 ===");

  var parent = getInspectionParentFolder();
  if (!parent) {
    Logger.log("エラー: 親フォルダ「" + INSPECTION_REPORT_CONFIG.FOLDER_PARENT + "」が見つかりません。FOLDER_PARENT_ID またはマイドライブ直下を確認してください。");
    return;
  }

  var inboxFolder = getInspectionInboxFolder(parent);
  if (!inboxFolder) {
    Logger.log("エラー: 受信フォルダが見つかりません。FOLDER_INBOX_ID を設定するか setupFoldersInspectionReport() を実行してください。");
    return;
  }

  var doneFolder = getInspectionDoneFolder(parent);
  if (!doneFolder) {
    doneFolder = parent.createFolder(INSPECTION_REPORT_CONFIG.FOLDER_DONE);
    Logger.log("処理済みフォルダを自動作成しました: " + INSPECTION_REPORT_CONFIG.FOLDER_DONE);
  }

  var files = inboxFolder.getFilesByType(MimeType.PDF);
  var pdfList = [];
  while (files.hasNext()) {
    pdfList.push(files.next());
  }

  if (pdfList.length === 0) {
    Logger.log("処理対象のPDFがありません。");
    return;
  }

  Logger.log("対象PDF: " + pdfList.length + " 件（受信・処理済みは同一親フォルダを参照）");

  var appData = getInspectionAppData();
  if (!appData) {
    Logger.log("エラー: 管理アプリのデータを取得できません。");
    return;
  }

  var results = [];
  pdfList.forEach(function(file) {
    Logger.log("\n--- 処理中: " + file.getName() + " ---");
    var result = processSingleInspectionPdf(file, appData, inboxFolder, doneFolder);
    if (result) {
      results.push(result);
    }
  });

  if (results.length > 0) {
    sendInspectionResultEmail(results);  // 異常があるときだけ送信（本文も異常のみ）
  }

  Logger.log("\n=== 処理完了: " + results.length + " 件 ===");
}

// ============================================================
// PDF処理
// ============================================================

function processSingleInspectionPdf(file, appData, inboxFolder, doneFolder) {
  var today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
  var originalName = file.getName();

  try {
    var text = extractTextFromInspectionPdf(file);
    if (!text) {
      movePdfToDoneFolder(file, "処理エラー_テキスト抽出失敗_" + today, originalName, inboxFolder, doneFolder);
      return { fileName: originalName, error: "テキスト抽出失敗" };
    }

    Logger.log("抽出テキスト（先頭500文字）:\n" + text.substring(0, 500));

    var storeName = extractStoreNameFromReport(text);
    if (!storeName) {
      movePdfToDoneFolder(file, "処理エラー_店舗名不明_" + today, originalName, inboxFolder, doneFolder);
      return { fileName: originalName, error: "店舗名を特定できません" };
    }
    Logger.log("店舗名: " + storeName);

    var reportCounts = extractCumulativeCountsFromReport(text);
    if (reportCounts.length === 0) {
      movePdfToDoneFolder(file, "点検報告書_" + storeName + "SS_累計台数不明_" + today, originalName, inboxFolder, doneFolder);
      return { fileName: originalName, storeName: storeName, error: "累計台数を読み取れません" };
    }
    Logger.log("報告書の累計台数: " + JSON.stringify(reportCounts));

    var reportDate = extractReportDate(text);  // 訪問日（なければ null）
    if (reportDate) {
      Logger.log("報告書の訪問日: " + Utilities.formatDate(reportDate, "Asia/Tokyo", "yyyy/MM/dd"));
    }

    var comparisons = compareInspectionWithAppData(storeName, reportCounts, appData, reportDate);

    var newName = "点検報告書_" + storeName + "SS_" + today + ".pdf";
    file.setName(newName);

    doneFolder.addFile(file);
    inboxFolder.removeFile(file);
    Logger.log("リネーム＆移動: " + newName + " → 処理済み");

    return {
      fileName: newName,
      storeName: storeName,
      comparisons: comparisons,
      reportCounts: reportCounts,
      extractedTextSnippet: text ? text.substring(0, 300) : null,
      pdfUrl: file.getUrl(),
      error: null
    };
  } catch (e) {
    Logger.log("エラー: " + e.toString());
    movePdfToDoneFolder(file, "処理エラー_" + today, originalName, inboxFolder, doneFolder);
    return { fileName: originalName, error: e.toString() };
  }
}

/**
 * エラー時でも受信フォルダに残さないよう、PDFを処理済みフォルダへ移動する
 * inboxFolder / doneFolder は processInspectionReports で取得した同じ参照を渡す
 */
function movePdfToDoneFolder(file, baseName, originalName, inboxFolder, doneFolder) {
  try {
    var safeName = baseName + ".pdf";
    if (safeName.length > 200) {
      safeName = baseName.substring(0, 196) + ".pdf";
    }
    file.setName(safeName);

    if (doneFolder) {
      doneFolder.addFile(file);
      if (inboxFolder) {
        inboxFolder.removeFile(file);
      }
      Logger.log("エラーのため処理済みへ移動: " + safeName);
    } else {
      var parent = getInspectionParentFolder();
      var done = parent ? getSubfolder(parent, INSPECTION_REPORT_CONFIG.FOLDER_DONE) : null;
      if (!done && parent) {
        done = parent.createFolder(INSPECTION_REPORT_CONFIG.FOLDER_DONE);
      }
      if (done) {
        done.addFile(file);
        var inbox = parent ? getSubfolder(parent, INSPECTION_REPORT_CONFIG.FOLDER_INBOX) : null;
        if (inbox) {
          inbox.removeFile(file);
        }
        Logger.log("エラーのため処理済みへ移動: " + safeName);
      } else {
        Logger.log("エラー: 処理済みフォルダを取得・作成できません。ファイルは受信フォルダに残っています。");
      }
    }
  } catch (moveErr) {
    Logger.log("移動エラー: " + moveErr.toString());
  }
}

/**
 * PDFからテキストを抽出する（Drive API で Google ドキュメントに変換して OCR）
 * ※ 拡張機能で Drive API を有効化してください。
 */
function extractTextFromInspectionPdf(pdfFile) {
  var parent = getInspectionParentFolder();
  if (!parent) {
    Logger.log("PDF変換エラー: 親フォルダ「" + INSPECTION_REPORT_CONFIG.FOLDER_PARENT + "」が見つかりません。");
    return null;
  }
  var tempFolder = getInspectionFolder(INSPECTION_REPORT_CONFIG.TEMP_FOLDER);
  if (!tempFolder) {
    tempFolder = parent.createFolder(INSPECTION_REPORT_CONFIG.TEMP_FOLDER);
  }

  try {
    var resource = {
      title: "temp_ocr_" + new Date().getTime(),
      mimeType: MimeType.GOOGLE_DOCS,
      parents: [{ id: tempFolder.getId() }]
    };

    // PDF→Google Doc 変換のみ。OCR は画像用のため PDF では使わない（指定するとエラーになる）
    // テキスト付きPDFは変換時に文字が取り込まれる。スキャンPDFのみの場合は文字が取れない場合あり
    var blob = pdfFile.getBlob();
    var docFile = Drive.Files.insert(resource, blob);

    var doc = DocumentApp.openById(docFile.id);
    var text = doc.getBody().getText();

    DriveApp.getFileById(docFile.id).setTrashed(true);

    return text;
  } catch (e) {
    Logger.log("PDF変換エラー: " + e.toString());
    return null;
  }
}

// ============================================================
// テキスト解析
// ============================================================

/**
 * 報告書テキストから訪問日を抽出する（例: 2026年02月04日）
 * 戻り値: Date（Asia/Tokyo）または null
 */
function extractReportDate(text) {
  if (!text) return null;
  // 訪問日の直後の日付（半角・全角数字対応）
  var match = text.match(/訪問日\s*([0-9０-９]{4})年([0-9０-９]{1,2})月([0-9０-９]{1,2})日/);
  if (!match) return null;
  var toHalf = function(s) {
    return s.replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  };
  var y = parseInt(toHalf(match[1]), 10);
  var m = parseInt(toHalf(match[2]), 10) - 1;
  var d = parseInt(toHalf(match[3]), 10);
  if (isNaN(y) || isNaN(m) || isNaN(d) || m < 0 || m > 11 || d < 1 || d > 31) return null;
  try {
    var date = new Date(y, m, d);
    if (date.getFullYear() !== y || date.getMonth() !== m || date.getDate() !== d) return null;
    return date;
  } catch (e) {
    return null;
  }
}

function extractStoreNameFromReport(text) {
  // 「セルフィックス◯◯SS」「セルフィックス◯◯給油所」など
  var patterns = [
    /セルフィックス(.+?)SS/,
    /セルフィックス(.+?)ＳＳ/,
    /セルフィックス(.+?)給油所/,
    /ｾﾙﾌｨｯｸｽ(.+?)SS/,
    /ｾﾙﾌｨｯｸｽ(.+?)給油所/
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match) {
      var rawName = match[1].trim();
      return normalizeInspectionStoreName(rawName);
    }
  }

  return null;
}

function normalizeInspectionStoreName(rawName) {
  var mapping = {
    "岡南": "岡南",
    "りんくう泉南": "りんくう",
    "りんくう": "りんくう",
    "池田": "池田",
    "小山": "小山",
    "天理": "天理",
    "厚木": "厚木",
    "裾野": "裾野",
    "倉吉": "倉吉",
    "糸我": "糸我",
    "貴志川": "貴志川",
    "かつらぎ": "かつらぎ",
    "和佐": "和佐",
    "紀三井寺": "紀三井寺",
    "和歌山北インター": "和歌山北インター",
    "東和歌山": "東和歌山",
    "御所": "御所",
    "熊野": "熊野",
    "坂出": "坂出",
    "徳島石井": "徳島石井",
    "小松島": "小松島",
    "牛久": "牛久",
    "土浦": "土浦",
    "岐阜東": "岐阜東",
    "太田": "太田",
    "北名古屋": "北名古屋",
    "ひたちなか": "ひたちなか"
  };

  if (mapping[rawName]) return mapping[rawName];

  var keys = Object.keys(mapping);
  for (var i = 0; i < keys.length; i++) {
    if (rawName.indexOf(keys[i]) >= 0 || keys[i].indexOf(rawName) >= 0) {
      return mapping[keys[i]];
    }
  }

  return rawName;
}

/**
 * テキストから累計洗車台数を抽出
 * 戻り値: [{position: "左", count: 51541}, ...]
 * 備考欄の「左記○○台、右記○○台」や「左機／右側」など書き方のゆらぎに対応
 */
function extractCumulativeCountsFromReport(text) {
  var results = [];
  var positionMap = {
    "左機": "左", "左": "左",
    "真ん中機": "中央", "真ん中": "中央", "中央機": "中央", "中央": "中央", "中": "中央",
    "右機": "右", "右": "右",
    "布機": "左"
  };

  function addOrUpdate(position, count) {
    var existing = results.find(function(r) { return r.position === position; });
    if (existing) {
      if (count > existing.count) existing.count = count;
    } else {
      results.push({ position: position, count: count });
    }
  }

  // 数字文字列をパース（全角数字・カンマ対応）
  function parseCountStr(str) {
    if (!str) return 0;
    var n = str.replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).replace(/,/g, "").replace(/\s/g, "").trim();
    var num = parseInt(n, 10);
    return isNaN(num) ? 0 : num;
  }

  // 半角・全角の数字にマッチする部分（キャプチャ用）
  var numPart = "([0-9０-９][0-9０-９,\\s]*)";

  var match;

  // パターン0: 「左側洗車機 13155台」「右側洗車機 14890台」「左側洗車機13155台」など（全角数字対応）
  var pattern0 = new RegExp("(左側洗車機|右側洗車機)\\s*" + numPart + "\\s*台", "g");
  while ((match = pattern0.exec(text)) !== null) {
    var count = parseCountStr(match[2]);
    if (count <= 0) continue;
    var position = match[1].indexOf("左") >= 0 ? "左" : "右";
    addOrUpdate(position, count);
  }

  // パターン1: 「左機 51541台」「真ん中機 58,624台」など
  var pattern1 = /(左機?|真ん中機?|中央機?|右機?|布機?)\s*[：:]?\s*(\d[\d,]*)\s*台?/g;
  while ((match = pattern1.exec(text)) !== null) {
    var posKey = match[1];
    var count = parseCountStr(match[2]);
    if (count <= 0) continue;
    var position = positionMap[posKey] || posKey;
    addOrUpdate(position, count);
  }

  // パターン1b: 備考欄「左記28436台」「左機→ 28436台」「右側→31133台」など（矢印・スペース・全角数字のゆらぎ対応）
  var arrowPart = "[\\s→⇒]*";  // 矢印やスペースを挟む場合
  var bikouPatterns = [
    { regex: new RegExp("(左記|左\\s*記|左機|左側|左)" + arrowPart + numPart + "\\s*台", "g"), position: "左" },
    { regex: new RegExp("(右記|右\\s*記|右機|右側|右)" + arrowPart + numPart + "\\s*台", "g"), position: "右" },
    { regex: new RegExp("(中央記|中央\\s*記|中央機|中央側|中央|真ん中)" + arrowPart + numPart + "\\s*台", "g"), position: "中央" }
  ];
  for (var b = 0; b < bikouPatterns.length; b++) {
    var bp = bikouPatterns[b];
    while ((match = bp.regex.exec(text)) !== null) {
      var count = parseCountStr(match[2]);
      if (count <= 0) continue;
      addOrUpdate(bp.position, count);
    }
  }

  if (results.length === 0) {
    var pattern2 = new RegExp("累計[洗車]*台数[：:\\s]*" + numPart, "g");
    while ((match = pattern2.exec(text)) !== null) {
      var count2 = parseCountStr(match[1]);
      if (count2 > 0) results.push({ position: "中央", count: count2 });
    }
  }

  if (results.length === 0) {
    var pattern3 = /累計[洗車]*台数[】\]]\s*([\s\S]{0,200})/;
    var block = text.match(pattern3);
    if (block) {
      var numPattern = /([0-9０-９][0-9０-９,\s]+)/g;
      var positions = ["左", "中央", "右"];
      var idx = 0;
      while ((match = numPattern.exec(block[1])) !== null && idx < 3) {
        var c = parseCountStr(match[1]);
        if (c > 0) {
          results.push({ position: positions[idx], count: c });
          idx++;
        }
      }
    }
  }

  return results;
}

// ============================================================
// データ比較（管理アプリ＝本システムのステータス集計）
// ============================================================

/**
 * アプリの累計台数の基準日（前月末日）を返す
 * 例: 本日が2月5日なら 1月31日
 */
function getInspectionAppReferenceDate() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), 0);  // 前月末
}

/**
 * 管理アプリのデータを取得（0_Config のスプレッドシート・ステータス集計を使用）
 * 戻り値: { "池田_右": {count, avg}, ... }
 */
function getInspectionAppData() {
  try {
    var config = getConfig();
    var ss = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(config.SHEET_NAMES.STATUS_SUMMARY);
    if (!sheet) {
      Logger.log("シート「" + config.SHEET_NAMES.STATUS_SUMMARY + "」が見つかりません");
      return null;
    }

    var data = sheet.getDataRange().getValues();
    var header = data[0];

    var colStore = header.indexOf("店舗名");
    var colPos = header.indexOf("区別");
    var colCount = header.indexOf("累計台数");
    var colAvg = header.indexOf("月平均台数");

    if (colStore < 0 || colPos < 0 || colCount < 0 || colAvg < 0) {
      Logger.log("必要な列が見つかりません: 店舗名/区別/累計台数/月平均台数");
      return null;
    }

    var result = {};
    for (var i = 1; i < data.length; i++) {
      var store = String(data[i][colStore]).trim();
      var pos = String(data[i][colPos]).trim();
      var count = parseInspectionNumber(data[i][colCount]);
      var avg = parseInspectionNumber(data[i][colAvg]);

      if (store && pos) {
        var key = store + "_" + pos;
        result[key] = { storeName: store, position: pos, count: count, avg: avg };
      }
    }

    Logger.log("管理アプリデータ: " + Object.keys(result).length + " 件取得");
    return result;
  } catch (e) {
    Logger.log("データ取得エラー: " + e.toString());
    return null;
  }
}

function parseInspectionNumber(val) {
  if (typeof val === "number") return val;
  var str = String(val).replace(/,/g, "").trim();
  var num = parseInt(str, 10);
  return isNaN(num) ? 0 : num;
}

/**
 * 店舗のアプリ側の位置（区別）一覧を取得
 */
function getAppPositionsForStore(storeName, appData) {
  var prefix = storeName + "_";
  var positions = [];
  for (var key in appData) {
    if (key.indexOf(prefix) === 0) {
      var pos = key.slice(prefix.length);
      if (pos) positions.push(pos);
    }
  }
  return positions;
}

/**
 * 報告書の複数値のうち、アプリ累計に最も近い値を返す（1台のみの店舗でパース誤りがある場合の補正用）
 */
function pickClosestReportCount(reportCounts, appCount) {
  if (!reportCounts || reportCounts.length === 0) return null;
  var closest = reportCounts[0];
  var minDiff = Math.abs(reportCounts[0].count - appCount);
  for (var i = 1; i < reportCounts.length; i++) {
    var d = Math.abs(reportCounts[i].count - appCount);
    if (d < minDiff) {
      minDiff = d;
      closest = reportCounts[i];
    }
  }
  return closest.count;
}

function compareInspectionWithAppData(storeName, reportCounts, appData, reportDate) {
  var comparisons = [];
  var thresholdMonths = INSPECTION_REPORT_CONFIG.THRESHOLD_MONTHS;
  var appRefDate = getInspectionAppReferenceDate();  // アプリの基準日＝前月末

  // アプリに存在する位置だけを比較する（報告書にない位置は無視、報告書だけにある位置も比較しない）
  var appPositions = getAppPositionsForStore(storeName, appData);
  if (appPositions.length === 0) {
    Logger.log(storeName + ": 管理アプリに該当店舗のデータがありません");
    return comparisons;
  }

  var singleMachineStore = (appPositions.length === 1);
  var reportCountByPos = {};
  reportCounts.forEach(function(r) { reportCountByPos[r.position] = r.count; });

  appPositions.forEach(function(appPosition) {
    var key = storeName + "_" + appPosition;
    var app = appData[key];
    if (!app) return;

    var reportCount;
    var reportHasThisPosition = reportCountByPos.hasOwnProperty(appPosition);
    var threshold = app.avg * thresholdMonths;  // 許容差（この範囲内なら✅）
    if (reportHasThisPosition) {
      reportCount = reportCountByPos[appPosition];
      // 1台のみの店舗で、報告のその位置の値が明らかにおかしい場合は、全報告値のうちアプリに最も近いものを採用。
      // ただし「最も近い値」でも許容差を超える場合は差し替えず、🔴としてメール通知する。
      if (singleMachineStore && reportCounts.length > 1) {
        var closestCount = pickClosestReportCount(reportCounts, app.count);
        var reportDiff = Math.abs(reportCount - app.count);
        var closestDiff = Math.abs(closestCount - app.count);
        if (closestDiff < reportDiff && closestDiff <= threshold) {
          reportCount = closestCount;
          Logger.log(storeName + " " + appPosition + ": 1台店のため報告値はアプリに最も近い数値を採用（" + reportCount + "）");
        }
      }
    } else {
      if (singleMachineStore && reportCounts.length >= 1) {
        reportCount = pickClosestReportCount(reportCounts, app.count);
        Logger.log(storeName + " " + appPosition + ": 報告書に該当表記なしのため、アプリに最も近い報告値を採用（" + reportCount + "）");
      } else {
        comparisons.push({
          position: appPosition,
          reportCount: null,
          appCount: app.count,
          predicted: null,
          diff: null,
          status: "⚠ 報告書に該当機の表記なし"
        });
        return;
      }
    }

    if (reportCount == null) return;

    // 予測値: アプリは前月末の累計。報告書の訪問日を考慮する
    var predicted;
    if (reportDate) {
      var refTime = new Date(appRefDate.getFullYear(), appRefDate.getMonth(), appRefDate.getDate()).getTime();
      var reportTime = new Date(reportDate.getFullYear(), reportDate.getMonth(), reportDate.getDate()).getTime();
      if (reportTime <= refTime) {
        predicted = app.count;
      } else {
        var daysDiff = (reportTime - refTime) / (24 * 60 * 60 * 1000);
        predicted = app.count + Math.round(app.avg * (daysDiff / 30));
      }
    } else {
      predicted = app.count + Math.round(app.avg * 1.5);
    }

    var diff = reportCount - predicted;

    var status;
    if (Math.abs(diff) <= threshold) {
      status = "✅ 正常";
    } else if (diff > 0) {
      status = "🔴 報告書の台数が多すぎる（+" + diff.toLocaleString() + "）";
    } else {
      if (reportCount < app.count) {
        status = "⚠ 報告書が古い可能性（報告＜アプリ）";
      } else {
        status = "🔴 報告書の台数が少なすぎる（" + diff.toLocaleString() + "）";
      }
    }

    comparisons.push({
      position: appPosition,
      reportCount: reportCount,
      appCount: app.count,
      appAvg: app.avg,
      predicted: predicted,
      diff: diff,
      status: status
    });

    Logger.log(storeName + " " + appPosition + ": " +
      "報告=" + reportCount + " / アプリ=" + app.count +
      " / 予測=" + predicted + " / 差=" + diff + " → " + status);
  });

  return comparisons;
}

// ============================================================
// 通知
// ============================================================

function sendInspectionResultEmail(results) {
  var body = "洗車機点検報告書の自動チェックで異常がありました。\n\n";
  body += "処理日時: " + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "\n";
  body += "━━━━━━━━━━━━━━━━━━━━━━\n\n";

  var hasAlert = false;

  results.forEach(function(result) {
    if (result.error) {
      body += "【エラー】" + result.fileName + "\n";
      body += "  " + result.error + "\n\n";
      hasAlert = true;
      return;
    }

    var abnormalComps = (result.comparisons || []).filter(function(comp) {
      return comp.status && comp.status.indexOf("🔴") >= 0;
    });
    if (abnormalComps.length === 0) return;

    hasAlert = true;
    body += "■ " + result.storeName + "SS（" + result.fileName + "）\n";
    if (result.pdfUrl) {
      body += "  PDF: " + result.pdfUrl + "\n";
    }
    abnormalComps.forEach(function(comp) {
      body += "  " + comp.position + "機: ";
      body += "報告=" + (comp.reportCount ? comp.reportCount.toLocaleString() : "?") + "台";
      if (comp.appCount !== null) {
        body += " / アプリ=" + comp.appCount.toLocaleString() + "台";
        body += " / 予測=" + comp.predicted.toLocaleString() + "台";
      }
      body += "\n  → " + comp.status + "\n";
    });
    // デバッグ情報: PDF読み取り結果の詳細
    if (result.reportCounts && result.reportCounts.length > 0) {
      body += "  [読取値] ";
      body += result.reportCounts.map(function(rc) {
        return rc.position + "=" + rc.count.toLocaleString();
      }).join(", ");
      body += "\n";
    }
    if (result.extractedTextSnippet) {
      body += "  [抽出テキスト先頭300字]\n  " + result.extractedTextSnippet.replace(/\n/g, "\n  ") + "\n";
    }
    body += "\n";
  });

  if (!hasAlert) {
    Logger.log("洗車機点検報告書チェック: 異常なしのためメール送信しません（結果件数=" + results.length + "）");
    return;
  }

  body += "━━━━━━━━━━━━━━━━━━━━━━\n";
  body += "\n上記を確認してください。\n";

  var to = INSPECTION_REPORT_CONFIG.NOTIFY_EMAIL;
  if (!to) {
    var config = getConfig();
    to = config.ADMIN_EMAIL || Session.getActiveUser().getEmail();
  }

  var subject = "【要確認】洗車機点検報告書チェックで異常あり";
  var fromEmail = INSPECTION_REPORT_CONFIG.FROM_EMAIL || Session.getActiveUser().getEmail();

  try {
    if (typeof Gmail !== 'undefined' && Gmail.Users && Gmail.Users.Messages) {
      var encodedSubject = encodeSubjectRFC2047(subject);
      var bodyBase64 = Utilities.base64Encode(Utilities.newBlob(body, 'UTF-8').getBytes());
      // 本文Base64を76文字ごとに改行（MIME規格準拠）
      var bodyLines = bodyBase64.match(/.{1,76}/g).join('\r\n');
      var mime = 'MIME-Version: 1.0\r\n'
        + 'Content-Type: text/plain; charset=UTF-8\r\n'
        + 'Content-Transfer-Encoding: base64\r\n'
        + 'From: ' + fromEmail + '\r\n'
        + 'To: ' + to + '\r\n'
        + 'Subject: ' + encodedSubject + '\r\n'
        + '\r\n'
        + bodyLines;
      var raw = Utilities.base64EncodeWebSafe(Utilities.newBlob(mime, 'text/plain').getBytes());
      Gmail.Users.Messages.send({ raw: raw }, 'me');
      Logger.log("洗車機点検報告書チェック: Gmail API で送信完了（From: " + fromEmail + " → " + to + "）");
      return;
    }
  } catch (e) {
    Logger.log("Gmail API 送信失敗、MailApp にフォールバック: " + e.toString());
  }

  MailApp.sendEmail(to, subject, body);
  Logger.log("洗車機点検報告書チェック: MailApp で通知メール送信完了 → " + to);
}

// ============================================================
// ユーティリティ
// ============================================================

/**
 * メール件名をRFC 2047 Base64エンコードする
 * 長い件名は複数のencoded-wordに分割（1語あたり75文字以内）
 */
function encodeSubjectRFC2047(subject) {
  var bytes = Utilities.newBlob(subject, 'UTF-8').getBytes();
  var prefix = '=?UTF-8?B?';
  var suffix = '?=';
  var wrapperLen = prefix.length + suffix.length;  // 12文字
  // 1 encoded-word あたりの Base64 部分の最大文字数（75 - wrapper = 63、ただし4の倍数に切り下げ）
  var maxBase64Chars = Math.floor((75 - wrapperLen) / 4) * 4;  // 60
  // Base64の60文字 = 元データ45バイト
  var maxBytesPerChunk = (maxBase64Chars / 4) * 3;  // 45

  var encoded = Utilities.base64Encode(bytes);
  if ((prefix + encoded + suffix).length <= 75) {
    return prefix + encoded + suffix;
  }

  // UTF-8の文字境界を壊さないようにバイト配列をチャンク分割してエンコード
  var parts = [];
  var i = 0;
  while (i < bytes.length) {
    var end = Math.min(i + maxBytesPerChunk, bytes.length);
    // UTF-8のマルチバイト文字の途中で切らない（先頭バイトの位置まで戻る）
    while (end < bytes.length && end > i && (bytes[end] & 0xC0) === 0x80) {
      end--;
    }
    var chunk = bytes.slice(i, end);
    parts.push(prefix + Utilities.base64Encode(chunk) + suffix);
    i = end;
  }
  return parts.join('\r\n ');
}

/** 親フォルダ「21_アピカ点検報告書」を取得（FOLDER_PARENT_ID が設定されていれば ID で、なければ名前検索） */
function getInspectionParentFolder() {
  var id = INSPECTION_REPORT_CONFIG.FOLDER_PARENT_ID;
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      Logger.log("親フォルダID取得エラー: " + e.toString());
      return null;
    }
  }
  var root = DriveApp.getRootFolder();
  var it = root.getFoldersByName(INSPECTION_REPORT_CONFIG.FOLDER_PARENT);
  return it.hasNext() ? it.next() : null;
}

/** 受信フォルダを取得（FOLDER_INBOX_ID が設定されていれば ID で、なければ親のサブフォルダ名で） */
function getInspectionInboxFolder(parent) {
  var id = INSPECTION_REPORT_CONFIG.FOLDER_INBOX_ID;
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      Logger.log("受信フォルダID取得エラー: " + e.toString());
      return null;
    }
  }
  return parent ? getSubfolder(parent, INSPECTION_REPORT_CONFIG.FOLDER_INBOX) : null;
}

/** 処理済みフォルダを取得（FOLDER_DONE_ID が設定されていれば ID で、なければ親のサブフォルダ名で） */
function getInspectionDoneFolder(parent) {
  var id = INSPECTION_REPORT_CONFIG.FOLDER_DONE_ID;
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      Logger.log("処理済みフォルダID取得エラー: " + e.toString());
      return null;
    }
  }
  return parent ? getSubfolder(parent, INSPECTION_REPORT_CONFIG.FOLDER_DONE) : null;
}

/** 指定した親フォルダ直下のサブフォルダを名前で取得（同じ親を渡せば一貫した参照になる） */
function getSubfolder(parentFolder, subfolderName) {
  if (!parentFolder) return null;
  var it = parentFolder.getFoldersByName(subfolderName);
  return it.hasNext() ? it.next() : null;
}

/** 親フォルダ直下のサブフォルダを名前で取得（getInspectionParentFolder を都度呼ぶため、同名親が複数あると別フォルダを指す可能性あり） */
function getInspectionFolder(subfolderName) {
  var parent = getInspectionParentFolder();
  return getSubfolder(parent, subfolderName);
}

function getFolderByName(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

/**
 * テスト用：手動で実行してフォルダ内のPDFを処理する
 */
function testProcessInspectionReports() {
  processInspectionReports();
}
