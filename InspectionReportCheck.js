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
 * 3. setupMonthlyTriggerInspectionReport() を実行して月次トリガーを設定
 *
 * 【運用】
 * 「アピカ点検報告書_受信」フォルダにPDFを入れると、processInspectionReports() 実行時に処理されます。
 */

// ============================================================
// 設定（点検報告書チェック専用）
// ============================================================
var INSPECTION_REPORT_CONFIG = {
  // Googleドライブのフォルダ名
  FOLDER_INBOX: "アピカ点検報告書_受信",
  FOLDER_DONE: "アピカ点検報告書_処理済み",
  TEMP_FOLDER: "アピカ点検_一時変換",

  // 通知先メールアドレス（未設定時は getConfig().ADMIN_EMAIL を使用）
  NOTIFY_EMAIL: null,

  // 累計台数のズレ許容範囲（月平均の何倍までOKか）
  THRESHOLD_MONTHS: 2
};

// ============================================================
// 初期セットアップ
// ============================================================

/**
 * 点検報告書用フォルダを自動作成する（最初に1回だけ実行）
 */
function setupFoldersInspectionReport() {
  var root = DriveApp.getRootFolder();
  var folders = [
    INSPECTION_REPORT_CONFIG.FOLDER_INBOX,
    INSPECTION_REPORT_CONFIG.FOLDER_DONE,
    INSPECTION_REPORT_CONFIG.TEMP_FOLDER
  ];

  folders.forEach(function(name) {
    var existing = DriveApp.getFoldersByName(name);
    if (existing.hasNext()) {
      Logger.log("既存フォルダ: " + name + " (ID: " + existing.next().getId() + ")");
    } else {
      var folder = root.createFolder(name);
      Logger.log("作成しました: " + name + " (ID: " + folder.getId() + ")");
    }
  });

  Logger.log("\n=== セットアップ完了 ===");
  Logger.log("「" + INSPECTION_REPORT_CONFIG.FOLDER_INBOX + "」フォルダにPDFを入れてください。");
}

/**
 * 点検報告書チェック用の月次トリガーを設定する（最初に1回だけ実行）
 * 毎月1日の午前9時に自動実行
 */
function setupMonthlyTriggerInspectionReport() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "processInspectionReports") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("processInspectionReports")
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log("点検報告書チェックの月次トリガーを設定しました（毎月1日 9:00-10:00）");
}

// ============================================================
// メイン処理
// ============================================================

/**
 * メイン関数：フォルダ内のPDFを処理する
 */
function processInspectionReports() {
  Logger.log("=== 点検報告書チェック開始 ===");

  var inboxFolder = getFolderByName(INSPECTION_REPORT_CONFIG.FOLDER_INBOX);
  if (!inboxFolder) {
    Logger.log("エラー: 受信フォルダが見つかりません。setupFoldersInspectionReport() を実行してください。");
    return;
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

  Logger.log("対象PDF: " + pdfList.length + " 件");

  var appData = getInspectionAppData();
  if (!appData) {
    Logger.log("エラー: 管理アプリのデータを取得できません。");
    return;
  }

  var results = [];
  pdfList.forEach(function(file) {
    Logger.log("\n--- 処理中: " + file.getName() + " ---");
    var result = processSingleInspectionPdf(file, appData);
    if (result) {
      results.push(result);
    }
  });

  if (results.length > 0) {
    sendInspectionResultEmail(results);
  }

  Logger.log("\n=== 処理完了: " + results.length + " 件 ===");
}

// ============================================================
// PDF処理
// ============================================================

function processSingleInspectionPdf(file, appData) {
  try {
    var text = extractTextFromInspectionPdf(file);
    if (!text) {
      return { fileName: file.getName(), error: "テキスト抽出失敗" };
    }

    Logger.log("抽出テキスト（先頭500文字）:\n" + text.substring(0, 500));

    var storeName = extractStoreNameFromReport(text);
    if (!storeName) {
      return { fileName: file.getName(), error: "店舗名を特定できません" };
    }
    Logger.log("店舗名: " + storeName);

    var reportCounts = extractCumulativeCountsFromReport(text);
    if (reportCounts.length === 0) {
      return { fileName: file.getName(), storeName: storeName, error: "累計台数を読み取れません" };
    }
    Logger.log("報告書の累計台数: " + JSON.stringify(reportCounts));

    var comparisons = compareInspectionWithAppData(storeName, reportCounts, appData);

    var today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
    var newName = "点検報告書_" + storeName + "SS_" + today + ".pdf";
    file.setName(newName);

    var doneFolder = getFolderByName(INSPECTION_REPORT_CONFIG.FOLDER_DONE);
    if (doneFolder) {
      doneFolder.addFile(file);
      var inbox = getFolderByName(INSPECTION_REPORT_CONFIG.FOLDER_INBOX);
      if (inbox) {
        inbox.removeFile(file);
      }
    }

    Logger.log("リネーム: " + newName);

    return {
      fileName: newName,
      storeName: storeName,
      comparisons: comparisons,
      error: null
    };
  } catch (e) {
    Logger.log("エラー: " + e.toString());
    return { fileName: file.getName(), error: e.toString() };
  }
}

/**
 * PDFからテキストを抽出する（Drive API で Google ドキュメントに変換して OCR）
 * ※ 拡張機能で Drive API を有効化してください。
 */
function extractTextFromInspectionPdf(pdfFile) {
  var tempFolder = getFolderByName(INSPECTION_REPORT_CONFIG.TEMP_FOLDER);
  if (!tempFolder) {
    tempFolder = DriveApp.getRootFolder().createFolder(INSPECTION_REPORT_CONFIG.TEMP_FOLDER);
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

function extractStoreNameFromReport(text) {
  var patterns = [
    /セルフィックス(.+?)SS/,
    /セルフィックス(.+?)ＳＳ/,
    /ｾﾙﾌｨｯｸｽ(.+?)SS/
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
 */
function extractCumulativeCountsFromReport(text) {
  var results = [];
  var positionMap = {
    "左機": "左", "左": "左",
    "真ん中機": "中央", "真ん中": "中央", "中央機": "中央", "中央": "中央", "中": "中央",
    "右機": "右", "右": "右",
    "布機": "左"
  };

  var pattern1 = /(左機?|真ん中機?|中央機?|右機?|布機?)\s*[：:]?\s*(\d[\d,]*)\s*台?/g;
  var match;
  while ((match = pattern1.exec(text)) !== null) {
    var posKey = match[1];
    var count = parseInt(match[2].replace(/,/g, ""), 10);
    var position = positionMap[posKey] || posKey;

    var existing = results.find(function(r) { return r.position === position; });
    if (existing) {
      if (count > existing.count) existing.count = count;
    } else {
      results.push({ position: position, count: count });
    }
  }

  if (results.length === 0) {
    var pattern2 = /累計[洗車]*台数[：:\s]*(\d[\d,]*)/g;
    while ((match = pattern2.exec(text)) !== null) {
      var count2 = parseInt(match[1].replace(/,/g, ""), 10);
      results.push({ position: "中央", count: count2 });
    }
  }

  if (results.length === 0) {
    var pattern3 = /累計[洗車]*台数[】\]]\s*([\s\S]{0,200})/;
    var block = text.match(pattern3);
    if (block) {
      var numPattern = /(\d[\d,]+)/g;
      var positions = ["左", "中央", "右"];
      var idx = 0;
      while ((match = numPattern.exec(block[1])) !== null && idx < 3) {
        results.push({
          position: positions[idx],
          count: parseInt(match[1].replace(/,/g, ""), 10)
        });
        idx++;
      }
    }
  }

  return results;
}

// ============================================================
// データ比較（管理アプリ＝本システムのステータス集計）
// ============================================================

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

function compareInspectionWithAppData(storeName, reportCounts, appData) {
  var comparisons = [];
  var thresholdMonths = INSPECTION_REPORT_CONFIG.THRESHOLD_MONTHS;

  reportCounts.forEach(function(report) {
    var key = storeName + "_" + report.position;
    var app = appData[key];

    if (!app) {
      comparisons.push({
        position: report.position,
        reportCount: report.count,
        appCount: null,
        predicted: null,
        diff: null,
        status: "⚠ 管理アプリにデータなし"
      });
      return;
    }

    var predicted = app.count + Math.round(app.avg * 1.5);
    var diff = report.count - predicted;
    var threshold = app.avg * thresholdMonths;

    var status;
    if (Math.abs(diff) <= threshold) {
      status = "✅ 正常";
    } else if (diff > 0) {
      status = "🔴 報告書の台数が多すぎる（+" + diff.toLocaleString() + "）";
    } else {
      status = "🔴 報告書の台数が少なすぎる（" + diff.toLocaleString() + "）";
    }

    comparisons.push({
      position: report.position,
      reportCount: report.count,
      appCount: app.count,
      appAvg: app.avg,
      predicted: predicted,
      diff: diff,
      status: status
    });

    Logger.log(storeName + " " + report.position + ": " +
      "報告=" + report.count + " / アプリ=" + app.count +
      " / 予測=" + predicted + " / 差=" + diff + " → " + status);
  });

  return comparisons;
}

// ============================================================
// 通知
// ============================================================

function sendInspectionResultEmail(results) {
  var hasAlert = false;
  var body = "洗車機点検報告書の自動チェック結果です。\n\n";
  body += "処理日時: " + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "\n";
  body += "処理件数: " + results.length + " 件\n";
  body += "━━━━━━━━━━━━━━━━━━━━━━\n\n";

  results.forEach(function(result) {
    if (result.error) {
      body += "【エラー】" + result.fileName + "\n";
      body += "  " + result.error + "\n\n";
      hasAlert = true;
      return;
    }

    body += "■ " + result.storeName + "SS（" + result.fileName + "）\n";

    result.comparisons.forEach(function(comp) {
      body += "  " + comp.position + "機: ";
      body += "報告=" + (comp.reportCount ? comp.reportCount.toLocaleString() : "?") + "台";

      if (comp.appCount !== null) {
        body += " / アプリ=" + comp.appCount.toLocaleString() + "台";
        body += " / 予測=" + comp.predicted.toLocaleString() + "台";
      }

      body += "\n  → " + comp.status + "\n";

      if (comp.status.indexOf("🔴") >= 0) {
        hasAlert = true;
      }
    });

    body += "\n";
  });

  body += "━━━━━━━━━━━━━━━━━━━━━━\n";

  if (hasAlert) {
    body += "\n⚠ アラートがあります。確認してください。\n";
  } else {
    body += "\n✅ すべて正常範囲内です。\n";
  }

  var subject = hasAlert
    ? "【要確認】洗車機点検報告書チェック結果"
    : "【正常】洗車機点検報告書チェック結果";

  var to = INSPECTION_REPORT_CONFIG.NOTIFY_EMAIL;
  if (!to) {
    var config = getConfig();
    to = config.ADMIN_EMAIL || Session.getActiveUser().getEmail();
  }

  MailApp.sendEmail(to, subject, body);
  Logger.log("通知メール送信完了: " + subject);
}

// ============================================================
// ユーティリティ
// ============================================================

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
