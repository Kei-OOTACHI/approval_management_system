/**
 * 文字列がURLかどうか確認する関数
 */
function isValidUrl_(url) {
  let pattern = /^(http|https):\/\/[^ "]+$/;
  return pattern.test(url);
}

/**
 * 文字列がメールアドレスかどうか確認する関数
 */
function isValidEmailAddress(email) {
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

/**
 * URLからファイルを取得する関数
 */
function getFileFromUrl_(url) {
  Logger.log("url:" + url);
  let parts = url.split('/');
  if (url.includes("viewform") || url.includes("forms.gle")) {
    return false;
  }
  let fileId = parts[parts.indexOf("d") + 1];
  Logger.log("fileId:" + fileId);
  let file = DriveApp.getFileById(fileId);
  return file;
}

/**
 * URLからフォルダを取得する関数
 */
function getFolderFromUrl_(folderUrl) {
  //ファイルurlと違って、フォルダurlはIDとクエリパラメータ(?~)が"/"で区切られていないので、?以下を切り落とす
  let match = folderUrl.match(/(?:\/folders\/|\/drive\/folders\/)([\w-]+)(?:\?.*)?/);
  Logger.log("folderId:" + match[1]);
  let folder = DriveApp.getFolderById(match[1]);
  return folder;
}

/**
 * アラートを表示する関数
 */
function setDialog_(title, prompt, caution = "") {
  const ui = SpreadsheetApp.getUi();
  if (caution == "") {
    ui.alert(title, prompt, ui.ButtonSet.OK);
  } else {
    const response = ui.alert(title, prompt, ui.ButtonSet.OK_CANCEL);
    if (response == ui.Button.OK) {
      return true;
    } else {
      ui.alert(caution, ui.ButtonSet.OK);
      return false;
    }
  }
}

/**
 * 入力欄のあるダイアログボックスを表示する関数
 * @parm {string} message - promptの文言
 * @return {string} 入力欄に入力した文字列
 */
function showInputDialog(message) {
  let ui = SpreadsheetApp.getUi();
  let result = ui.prompt(message, ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() === ui.Button.OK) {
    let imputValue = result.getResponseText();
    return imputValue;
  }
}

/**
 * Google Chatにメッセージを自動送信する関数
 * @parm {string} message - 送信文
 * @parm {string} webhookUrl - 送信先スペースのwebhook URL
 */
function ChatApp(message, webhookUrl) {
  let payload = {
    'text': message
  };

  let options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let response = UrlFetchApp.fetch(webhookUrl, options);
  Logger.log(response.getContentText());
}

