// 本体。トリガーで編集時着火
function whenEdit_(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (["設定", "このSSの使い方"].includes(sheetName)) return;

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  Logger.log(`${sheetName} シートの ${row} 行 ${col} 列が編集されました`);
  Logger.log(rowData);

  const team = rowData[COL_ROLE.team - 1];
  const url = rowData[COL_ROLE.url - 1];
  const deliberations = rowData.slice(COL_ROLE.del[0] - 1, COL_ROLE.del[COL_ROLE.del.length - 1]);

  if (col == COL_ROLE.team && team) {
    handleTeamEdit_(sheet, row, team, url);
  } else if (col == COL_ROLE.url && url) {
    handleUrlEdit_(e, sheet, row, url, team);
  } else if (COL_ROLE.del.includes(col) && !deliberations.includes(DEL_BEFORE[0])) {
    handleDeliberationEdit_(sheet, row, deliberations, rowData, url);
  }
}

// team列が編集されたときの処理
function handleTeamEdit_(sheet, row, team, url) {
  fillDeliberation_(sheet, row, team);
  if (url) {
    noticeSubmission_(sheet, row, team);
  }
}

// url列が編集されたときの処理
function handleUrlEdit_(e, sheet, row, url, team) {
  const file = getFileFromUrl_(url);
  if (!file) {
    setDialog_("URLエラー", "フォームは編集者用のリンクを挿入してください。");
    return;
  }

  changeAuthAndLocation_(file);
  fillTitleEditorDate_(e, sheet, row, file);

  if (team) {
    fillDeliberation_(sheet, row, team);
    noticeSubmission_(sheet, row, team);
  } else {
    setDialog_("注意～！", "チーム名を記入するのを忘れないでね～");
  }
}

// del(deliberation)列が編集されたときの処理
function handleDeliberationEdit_(sheet, row, deliberations, rowData, url) {
  const wholeDel = checkDeliberation_(sheet, row, deliberations);
  if (wholeDel && showConfirmDialog_()) {
    const editorEmail = rowData[COL_ROLE.editorEmail - 1];
    if (checkUrlAndEmail_(url, editorEmail)) {
      noticeDelResult_(sheet, row, wholeDel, editorEmail);
      saveFile_(sheet, row, rowData, url);
    }
  }
}

// 審議内容を埋め込む
function fillDeliberation_(sheet, row, team) {
  let teamNum = TEAM_OPTIONS.indexOf(team);
  let teamSupervisors = SUPERVISOR_OPTIONS[teamNum].slice(1);
  let orgDel = teamSupervisors.map(value => value ? DEL_BEFORE[0] : DEL_BEFORE[1]);
  sheet.getRange(row, COL_ROLE.del[0], 1, COL_ROLE.del.length).setValues([orgDel]);
}

// ファイルの権限と場所を変更
function changeAuthAndLocation_(file) {
  // メールで通知を送信しないで編集権限付与
  if (SUPERVISOR_ADRESS.length) {
    try {
      Drive.Permissions.insert({
        'role': 'writer',
        'type': 'user',
        'value': SUPERVISOR_ADRESS.filter(Boolean)
      }, file.getId(), {
        'sendNotificationEmails': false
      });
    } catch (e) {
      Logger.log(e);
      setDialog_("編集権限自動付与失敗", e + "\n手動でファイルの編集権限を敲き担当に渡してください");
    }
  }
  // フォルダの移動
  if (FOLDER_URL) {
    let folder = getFolderFromUrl_(FOLDER_URL);
    try {
      file.moveTo(folder);
    } catch (e) {
      Logger.log(e);
      setDialog_("ファイル移動失敗", e + "\n手動でファイルを指定されたフォルダに移動してください");
    }
  }
}

// ファイル情報をシートに埋め込む
function fillTitleEditorDate_(e, sheet, row, file) {
  const fileName = file.getName();
  const editorEmail = e.user.getEmail();
  const editDate = new Date();
  sheet.getRange(row, COL_ROLE.fileName).setValue(fileName);
  sheet.getRange(row, COL_ROLE.editorEmail).setValue(editorEmail);
  sheet.getRange(row, COL_ROLE.editDate).setValue(editDate);
}

// 提出通知を送信
function noticeSubmission_(sheet, row, team) {
  let message = makeMessage_(sheet, row, DEL_BEFORE[0]);
  message = addMention_(message, team);
  ChatApp(message, WEBHOOK_URL);
  setDialog_("通知完了", "");
}

// 審議結果に合わせて送信内容を選択＆送信前確認
function checkDeliberation_(sheet, row, deliberations) {
  let wholeDel = DEL_AFTER.find(option => deliberations.includes(option));
  if (wholeDel == "その他") {
    setDialog_("確認", makeMessage_(sheet, row, wholeDel));
    return false;
  }
  return wholeDel;
}

// 送信前の確認ダイアログを表示
function showConfirmDialog_() {
  const confirm = setDialog_(
    "送信前確認",
    "ドキュメントを保存し、提出者に敲きが終わったことを通知します。よろしいですか？",
    'キャンセルしました\n手動でドキュメントを保存資料保管スペースに保存し、提出者に敲きが終わったことを連絡してください'
  );
  return confirm;
}

// URLとメールアドレスの有効性をチェック
function checkUrlAndEmail_(url, editorEmail) {
  if (!isValidUrl_(url)) {
    setDialog_("URLエラー", "C列のURLが無効です。スマートチップではなくURLのままで提出してください。\n今回自動通知は実行されないので、手動で通知をお願いします。");
    return false;
  }
  if (!isValidEmailAddress(editorEmail)) {
    setDialog_("メールアドレスエラー", "E列の提出者メールアドレスが無効です。スマートチップではなくメールアドレスのままで提出してください。\n今回自動通知は実行されないので、手動で通知をお願いします。");
    return false;
  }
  return true;
}

// 提出結果を通知
function noticeDelResult_(sheet, row, wholeDel, editorEmail) {
  let message = makeMessage_(sheet, row, wholeDel);
  sendDirectMessage(editorEmail, message);
  setDialog_("送信完了", "");
}

// ファイル情報を保存
function saveFile_(sheet, row, rowData, url) {
  const saveCol = searchLastColumnWithData_(rowData) + 1;
  let submitDate = rowData[COL_ROLE.editDate - 1];
  submitDate = dateFormatter_(submitDate);
  const recode = SpreadsheetApp.newRichTextValue()
    .setText(submitDate)
    .setLinkUrl(url)
    .build();
  Logger.log(`recode${recode}`);
  sheet.getRange(row, saveCol).setRichTextValue(recode);
}

// 日付をフォーマット
function dateFormatter_(date) {
  return date instanceof Date ? Utilities.formatDate(date, 'JST', 'M/d') : setDialog_("エラー", "提出日を入力してください");
}

// 最後にデータが入力された列を検索
function searchLastColumnWithData_(rowData) {
  for (let i = rowData.length - 1; i >= COL_ROLE.firstDraft - 2; i--) {
    if (rowData[i] !== "") return i + 1;
  }
  return COL_ROLE.firstDraft;
}

// メッセージテンプレートを作成
function makeMessage_(sheet, row, deliberation) {
  const msTempNum = MESSAGE_OPTIONS[0].indexOf(deliberation);
  let message = MESSAGE_OPTIONS[2][msTempNum];
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  let submitDate = dateFormatter_(rowData[COL_ROLE.editDate - 1]);
  message = message
    .replace(/\\team\\/g, rowData[COL_ROLE.team - 1])
    .replace(/\\url\\/g, rowData[COL_ROLE.url - 1])
    .replace(/\\fileName\\/g, rowData[COL_ROLE.fileName - 1])
    .replace(/\\editorEmail\\/g, rowData[COL_ROLE.editorEmail - 1])
    .replace(/\\submitDate\\/g, submitDate)
    .replace(/\\sheetName\\/g, sheet.getName());
  return message;
}

// 送信メッセージにメンション部分を追加
function addMention_(message, team) {
  let teamNum = TEAM_OPTIONS.indexOf(team);//チームが上から数えて何番目なのか
  let teamSupervisors = SUPERVISOR_OPTIONS[teamNum].slice(1);//そのチームの敲き担当者が誰なのか。bool値の配列

  let mentionedMessage = message;
  for (i = 0; i < teamSupervisors.length; i++) {//この反復処理で敲き担当者をメンションする文字列を文面に追加。
    if (teamSupervisors[i]) {
      var mentionTo = SUPERVISOR_ADRESS[i];
      if (!mentionTo) continue;
      var user = AdminDirectory.Users.get(mentionTo, { viewType: 'domain_public' });
      mentionedMessage = "<users/" + user.id + ">\n" + mentionedMessage;
    }
  }

  return mentionedMessage;
}

// トリガーを設定
function setTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('whenEdit_').forSpreadsheet(ss).onEdit().create();
}

// テスト用関数
function test() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getActiveSheet().getActiveRange();
  const mail = "kei.ootachi@wasedasai.net";
  const dummyE = { source: ss, range: range, user: { getEmail: mail } };
  whenEdit_(dummyE);
}

// スプレッドシートが開かれたときに実行される関数
function onOpen() {
  setDialog_("連絡", "この敲きFM（自動入力、自動通知システム付き🤩）を使用したい場合は、DXPJ員かkei.ootachi@wasedasai.netまでご連絡ください～\n自動化のためのもろもろの初期設定をしますので!");
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("管理者用設定").addItem("⚠管理者のみ押下/一度権限を付与したら再度押下⚠トリガーを設置", "setTrigger_").addToUi();
}
