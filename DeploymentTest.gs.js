var APP_VERSION = "0.1.1";

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("開発管理")
    .addItem("動作環境を確認", "writeDeploymentTestLog")
    .addToUi();
}
/**
 * 現在動いているGAS・スプレッドシートの情報をログに記録する。
 */
function writeDeploymentTestLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var environment = getConfigValue_(ss, "ENV") || "未設定";
  var githubBranch = getConfigValue_(ss, "GITHUB_BRANCH") || "未設定";

  var sheetName = "DEPLOY_TEST_LOG";
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);

    sheet.appendRow([
      "executedAt",
      "environment",
      "githubBranch",
      "appVersion",
      "spreadsheetName",
      "spreadsheetId",
      "scriptId"
    ]);

    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    environment,
    githubBranch,
    APP_VERSION,
    ss.getName(),
    ss.getId(),
    ScriptApp.getScriptId()
  ]);

  ss.toast(
    "環境: " + environment +
      "\nバージョン: " + APP_VERSION,
    "動作確認完了",
    5
  );

  Logger.log("動作確認ログを記録しました。");
  Logger.log("environment: " + environment);
  Logger.log("githubBranch: " + githubBranch);
  Logger.log("appVersion: " + APP_VERSION);
}

