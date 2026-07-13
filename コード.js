function getFolderOfId(spreadsheetId) {
  if (!spreadsheetId) {
    Logger.log("エラー: スプレッドシートIDが空です。");
    return null;
  }

  
    var file = DriveApp.getFileById(spreadsheetId);
    var pp = file.getParents();
    
    if (pp.hasNext()) {
      return pp.next();
    } else {
      Logger.log("親フォルダが見つかりませんでした（マイドライブ直下の可能性があります）。");
      return null;
    }
}
  
function setupFeatureEnvironment() {
  var featureName = "test-feature-1"; // まずはここを書き換えて使う
  return createFeatureEnvironment(featureName);
}


function createFeatureEnvironment(featureName) {
  if (!featureName) {
    throw new Error("featureName が空です。");
  }

  var masterSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSpreadsheetId = masterSpreadsheet.getId();
  var masterSpreadsheetName = masterSpreadsheet.getName();

  // feature用GASからさらにコピーしてしまう事故を防ぐ
  var env = getConfigValue_(masterSpreadsheet, "ENV");
  if (env === "feature") {
    throw new Error("feature環境からさらにfeature環境を作成することは想定していません。masterから実行してください。");
  }

  var masterFile = DriveApp.getFileById(masterSpreadsheetId);
  var parentFolder = getFolderOfId(masterSpreadsheetId);

  var branchName = "feature-" + toSafeBranchName_(featureName);

  var timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss"
  );

  var copyName =
    masterSpreadsheetName +
    "_feature_" +
    toSafeFileName_(featureName) +
    "_" +
    timestamp;

  var copiedFile;
  if (parentFolder) {
    copiedFile = masterFile.makeCopy(copyName, parentFolder);
  } else {
    copiedFile = masterFile.makeCopy(copyName);
  }

  var featureSpreadsheetId = copiedFile.getId();
  var featureSpreadsheetUrl = copiedFile.getUrl();
  var featureSpreadsheet = SpreadsheetApp.openById(featureSpreadsheetId);

  writeFeatureConfig_(
    featureSpreadsheet,
    featureName,
    branchName,
    masterSpreadsheetId,
    masterSpreadsheetName,
    featureSpreadsheetId,
    copyName
  );

  appendEnvironmentRecord_(
    masterSpreadsheet,
    featureName,
    branchName,
    masterSpreadsheetId,
    masterSpreadsheetName,
    featureSpreadsheetId,
    copyName,
    featureSpreadsheetUrl
  );

  Logger.log("feature環境を作成しました。");
  Logger.log("featureName: " + featureName);
  Logger.log("branchName: " + branchName);
  Logger.log("spreadsheetName: " + copyName);
  Logger.log("spreadsheetUrl: " + featureSpreadsheetUrl);

  return {
    featureName: featureName,
    branchName: branchName,
    spreadsheetId: featureSpreadsheetId,
    spreadsheetName: copyName,
    spreadsheetUrl: featureSpreadsheetUrl
  };
}


function writeFeatureConfig_(
  spreadsheet,
  featureName,
  branchName,
  masterSpreadsheetId,
  masterSpreadsheetName,
  featureSpreadsheetId,
  featureSpreadsheetName
) {
  var sheet = spreadsheet.getSheetByName("CONFIG");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("CONFIG");
  }

  setConfigValue_(sheet, "ENV", "feature");
  setConfigValue_(sheet, "FEATURE_NAME", featureName);
  setConfigValue_(sheet, "GITHUB_BRANCH", branchName);
  setConfigValue_(sheet, "MASTER_SPREADSHEET_ID", masterSpreadsheetId);
  setConfigValue_(sheet, "MASTER_SPREADSHEET_NAME", masterSpreadsheetName);
  setConfigValue_(sheet, "FEATURE_SPREADSHEET_ID", featureSpreadsheetId);
  setConfigValue_(sheet, "FEATURE_SPREADSHEET_NAME", featureSpreadsheetName);
  setConfigValue_(sheet, "CREATED_AT", new Date());
  // setConfigValue_(sheet, "CREATED_BY", Session.getActiveUser().getEmail());
}


function appendEnvironmentRecord_(
  masterSpreadsheet,
  featureName,
  branchName,
  masterSpreadsheetId,
  masterSpreadsheetName,
  featureSpreadsheetId,
  featureSpreadsheetName,
  featureSpreadsheetUrl
) {
  var sheet = masterSpreadsheet.getSheetByName("ENVIRONMENTS");
  if (!sheet) {
    sheet = masterSpreadsheet.insertSheet("ENVIRONMENTS");
    sheet.appendRow([
      "createdAt",
      "status",
      "featureName",
      "branchName",
      "masterSpreadsheetName",
      "masterSpreadsheetId",
      "featureSpreadsheetName",
      "featureSpreadsheetId",
      "featureSpreadsheetUrl",
      "createdBy"
    ]);
  }

  sheet.appendRow([
    new Date(),
    "working",
    featureName,
    branchName,
    masterSpreadsheetName,
    masterSpreadsheetId,
    featureSpreadsheetName,
    featureSpreadsheetId,
    featureSpreadsheetUrl,
    // Session.getActiveUser().getEmail()
  ]);
}


function setConfigValue_(sheet, key, value) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["key", "value"]);
  }

  var lastRow = sheet.getLastRow();

  // ヘッダーがない場合に備える
  var firstRow = sheet.getRange(1, 1, 1, 2).getValues()[0];
  if (firstRow[0] !== "key" || firstRow[1] !== "value") {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, 2).setValues([["key", "value"]]);
    lastRow = sheet.getLastRow();
  }

  var values = sheet.getRange(1, 1, lastRow, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  sheet.appendRow([key, value]);
}


function getConfigValue_(spreadsheet, key) {
  var sheet = spreadsheet.getSheetByName("CONFIG");
  if (!sheet) {
    return null;
  }

  var values = sheet.getDataRange().getValues();

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      return values[i][1];
    }
  }

  return null;
}


function toSafeBranchName_(name) {
  return String(name)
    .trim()
    .toLowerCase()
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "");
}


function toSafeFileName_(name) {
  return String(name)
    .trim()
    .replace(/[\\/:*?"<>|#%{}]/g, "-")
    .replace(/\s+/g, "-");
}

function setupFeatureEnvironmentAndCreateBranch() {
  var featureName = "test-feature-3"; // ここを書き換えて使う

  var env = createFeatureEnvironment(featureName);

  createGitHubBranchIfNotExists_(env.branchName);

  markEnvironmentStatus_(env.branchName, "branch_created");

  upsertEnvironmentsMd_("main");

  Logger.log("Step 3 完了: ENVIRONMENTS.mdをGitHubに反映しました。");
  Logger.log("branchName: " + env.branchName);
  Logger.log("spreadsheetUrl: " + env.spreadsheetUrl);

  return env;
}


function createGitHubBranchIfNotExists_(branchName) {
  if (!branchName) {
    throw new Error("branchName が空です。");
  }

  var props = PropertiesService.getScriptProperties();
  var baseBranch = props.getProperty("GITHUB_BASE_BRANCH") || "main";

  var targetRefPath = "/git/ref/heads/" + branchName;
  var existingRef = githubRequest_("get", targetRefPath, null, true);

  if (existingRef) {
    Logger.log("GitHub branch は既に存在します: " + branchName);
    return existingRef;
  }

  var baseRefPath = "/git/ref/heads/" + baseBranch;
  var baseRef = githubRequest_("get", baseRefPath, null, false);

  if (!baseRef || !baseRef.object || !baseRef.object.sha) {
    throw new Error("base branch のSHAを取得できませんでした: " + baseBranch);
  }

  var baseSha = baseRef.object.sha;

  var payload = {
    ref: "refs/heads/" + branchName,
    sha: baseSha
  };

  var createdRef = githubRequest_("post", "/git/refs", payload, false);

  Logger.log("GitHub branch を作成しました: " + branchName);
  Logger.log("baseBranch: " + baseBranch);
  Logger.log("baseSha: " + baseSha);

  return createdRef;
}


function githubRequest_(method, path, payload, allow404) {
  var props = PropertiesService.getScriptProperties();

  var owner = props.getProperty("GITHUB_OWNER");
  var repo = props.getProperty("GITHUB_REPO");
  var token = props.getProperty("GITHUB_TOKEN");

  if (!owner) {
    throw new Error("GITHUB_OWNER が設定されていません。");
  }
  if (!repo) {
    throw new Error("GITHUB_REPO が設定されていません。");
  }
  if (!token) {
    throw new Error("GITHUB_TOKEN が設定されていません。");
  }

  var url = "https://api.github.com/repos/" + owner + "/" + repo + path;

  var options = {
    method: method,
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json"
    },
    muteHttpExceptions: true
  };

  if (payload !== null && payload !== undefined) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(payload);
  }

  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();
  var responseText = response.getContentText();

  if (statusCode === 404 && allow404) {
    return null;
  }

  if (statusCode >= 200 && statusCode < 300) {
    if (!responseText) {
      return {};
    }
    return JSON.parse(responseText);
  }

  throw new Error(
    "GitHub API Error\n" +
    "method: " + method + "\n" +
    "path: " + path + "\n" +
    "status: " + statusCode + "\n" +
    "body: " + responseText
  );
}


function markEnvironmentStatus_(branchName, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ENVIRONMENTS");

  if (!sheet) {
    Logger.log("ENVIRONMENTS シートがないため、status更新をスキップしました。");
    return;
  }

  var values = sheet.getDataRange().getValues();

  if (values.length === 0) {
    return;
  }

  var headers = values[0];
  var branchIndex = headers.indexOf("branchName");
  var statusIndex = headers.indexOf("status");

  if (branchIndex === -1 || statusIndex === -1) {
    Logger.log("ENVIRONMENTS シートに branchName または status 列がありません。");
    return;
  }

  for (var i = 1; i < values.length; i++) {
    if (values[i][branchIndex] === branchName) {
      sheet.getRange(i + 1, statusIndex + 1).setValue(status);
      return;
    }
  }

  Logger.log("対象branchがENVIRONMENTSに見つかりませんでした: " + branchName);
}

function upsertEnvironmentsMd_(branchName) {
  if (!branchName) {
    throw new Error("branchName が空です。");
  }

  var content = buildEnvironmentsMarkdown_();

  var existingFile = getGitHubFileIfExists_("ENVIRONMENTS.md", branchName);

  var payload = {
    message: "Update ENVIRONMENTS.md",
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: branchName
  };

  if (existingFile && existingFile.sha) {
    payload.sha = existingFile.sha;
  }

  var result = githubRequest_(
    "put",
    "/contents/ENVIRONMENTS.md",
    payload,
    false
  );

  Logger.log("ENVIRONMENTS.md を更新しました。");
  Logger.log("branch: " + branchName);
  Logger.log("commit sha: " + result.commit.sha);

  return result;
}


function getGitHubFileIfExists_(filePath, branchName) {
  var path =
    "/contents/" +
    encodeURIComponent(filePath) +
    "?ref=" +
    encodeURIComponent(branchName);

  return githubRequest_("get", path, null, true);
}


function buildEnvironmentsMarkdown_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ENVIRONMENTS");

  if (!sheet) {
    throw new Error("ENVIRONMENTS シートが見つかりません。");
  }

  var values = sheet.getDataRange().getValues();

  if (values.length === 0) {
    throw new Error("ENVIRONMENTS シートが空です。");
  }

  var lines = [];

  lines.push("# Environments");
  lines.push("");
  lines.push("このファイルは master スプレッドシートの `ENVIRONMENTS` シートから自動生成しています。");
  lines.push("");
  lines.push("| createdAt | status | featureName | branchName | masterSpreadsheetName | featureSpreadsheetName | featureSpreadsheetUrl | createdBy |");
  lines.push("|---|---|---|---|---|---|---|---|");

  var headers = values[0];

  var createdAtIndex = headers.indexOf("createdAt");
  var statusIndex = headers.indexOf("status");
  var featureNameIndex = headers.indexOf("featureName");
  var branchNameIndex = headers.indexOf("branchName");
  var masterSpreadsheetNameIndex = headers.indexOf("masterSpreadsheetName");
  var featureSpreadsheetNameIndex = headers.indexOf("featureSpreadsheetName");
  var featureSpreadsheetUrlIndex = headers.indexOf("featureSpreadsheetUrl");
  var createdByIndex = headers.indexOf("createdBy");

  for (var i = 1; i < values.length; i++) {
    var row = values[i];

    if (!row[featureNameIndex] && !row[branchNameIndex]) {
      continue;
    }

    lines.push(
      "|" +
        [
          formatMarkdownCell_(row[createdAtIndex]),
          formatMarkdownCell_(row[statusIndex]),
          formatMarkdownCell_(row[featureNameIndex]),
          formatMarkdownCell_(row[branchNameIndex]),
          formatMarkdownCell_(row[masterSpreadsheetNameIndex]),
          formatMarkdownCell_(row[featureSpreadsheetNameIndex]),
          formatMarkdownCell_(row[featureSpreadsheetUrlIndex]),
          formatMarkdownCell_(row[createdByIndex])
        ].join("|") +
        "|"
    );
  }

  lines.push("");
  lines.push("Generated at: " + new Date().toISOString());
  lines.push("");

  return lines.join("\n");
}


function formatMarkdownCell_(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (Object.prototype.toString.call(value) === "[object Date]") {
    value = Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );
  }

  return String(value)
    .replace(/\|/g, "\\|")
    .replace(/\r?\n/g, "<br>");
}

/**
 * Step 4:
 * 現在開いている feature GAS プロジェクトの Code.gs / appsscript.json 等を
 * CONFIG の GITHUB_BRANCH に commit する。
 *
 * これは feature用スプレッドシート側のGASで実行する。
 */
function commitCurrentGasProjectToGitHubFeatureBranch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var env = getConfigValue_(ss, "ENV");
  if (env !== "feature") {
    throw new Error("この関数は feature環境のGASから実行してください。現在のENV: " + env);
  }

  var branchName = getConfigValue_(ss, "GITHUB_BRANCH");
  var featureName = getConfigValue_(ss, "FEATURE_NAME");

  if (!branchName) {
    throw new Error("CONFIG の GITHUB_BRANCH が空です。");
  }

  var files = getCurrentAppsScriptFiles_();

  if (!files || files.length === 0) {
    throw new Error("Apps Scriptファイルを取得できませんでした。");
  }

  Logger.log("commit対象branch: " + branchName);
  Logger.log("取得ファイル数: " + files.length);

  var results = [];

  files.forEach(function(file) {
    var repoPath = appsScriptFileToRepoPath_(file);

    // 念のため、変換できないファイルはスキップ
    if (!repoPath) {
      Logger.log("skip: " + JSON.stringify(file));
      return;
    }

    var message =
      "Update " +
      repoPath +
      " from GAS feature environment" +
      (featureName ? " (" + featureName + ")" : "");

    var result = upsertGitHubFile_(
      repoPath,
      file.source || "",
      branchName,
      message
    );

    results.push({
      path: repoPath,
      commitSha: result && result.commit ? result.commit.sha : null
    });

    Logger.log("committed: " + repoPath);
  });

  Logger.log("Step 4 完了: feature GAS のコードをGitHub branchへ反映しました。");
  Logger.log(JSON.stringify(results, null, 2));

  return results;
}


/**
 * Apps Script APIで、現在のGASプロジェクトのファイル一覧を取得する。
 */
function getCurrentAppsScriptFiles_() {
  var scriptId = ScriptApp.getScriptId();

  if (!scriptId) {
    throw new Error("Script ID を取得できませんでした。");
  }

  var path = "/v1/projects/" + encodeURIComponent(scriptId) + "/content";
  var content = googleAppsScriptApiRequest_("get", path, null);

  if (!content || !content.files) {
    throw new Error("Apps Script API のレスポンスに files がありません。");
  }

  return content.files;
}


/**
 * Google Apps Script API用のリクエスト関数。
 * ScriptApp.getOAuthToken() を使って、自分のGoogleアカウント権限でAPIを叩く。
 */
function googleAppsScriptApiRequest_(method, path, payload) {
  var url = "https://script.googleapis.com" + path;

  var options = {
    method: method,
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    }
  };

  if (payload !== null && payload !== undefined) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(payload);
  }

  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();
  var responseText = response.getContentText();

  if (statusCode >= 200 && statusCode < 300) {
    if (!responseText) {
      return {};
    }
    return JSON.parse(responseText);
  }

  throw new Error(
    "Apps Script API Error\n" +
    "method: " + method + "\n" +
    "path: " + path + "\n" +
    "status: " + statusCode + "\n" +
    "body: " + responseText
  );
}


/**
 * Apps Script API上のファイル情報を、GitHub上のファイルパスに変換する。
 *
 * Apps Script APIでは、
 * Code.gs       → name: "Code", type: "SERVER_JS"
 * appsscript.json → name: "appsscript", type: "JSON"
 * index.html    → name: "index", type: "HTML"
 * のように返る。
 */
function appsScriptFileToRepoPath_(file) {
  if (!file || !file.name || !file.type) {
    return null;
  }

  if (file.type === "SERVER_JS") {
    return file.name + ".gs";
  }

  if (file.type === "HTML") {
    return file.name + ".html";
  }

  if (file.type === "JSON") {
    return file.name + ".json";
  }

  return null;
}


/**
 * GitHub上の指定branchに、1ファイルを作成・更新する。
 */
function upsertGitHubFile_(filePath, content, branchName, commitMessage) {
  var existingFile = getGitHubFileIfExists_(filePath, branchName);

  var payload = {
    message: commitMessage,
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: branchName
  };

  if (existingFile && existingFile.sha) {
    payload.sha = existingFile.sha;
  }

  var result = githubRequest_(
    "put",
    "/contents/" + encodeGitHubContentPath_(filePath),
    payload,
    false
  );

  return result;
}


/**
 * GitHub Contents API用に path をエンコードする。
 * 将来 src/Code.gs のようなパスにしても壊れにくいように segment ごとに処理する。
 */
function encodeGitHubContentPath_(filePath) {
  return filePath
    .split("/")
    .map(function(part) {
      return encodeURIComponent(part);
    })
    .join("/");
}
