// 列の追加・変更があった場合は以下のグローバル変数を操作する
const COL_ROLE = {
  team: 1,
  url: 3,
  fileName: 4,
  editorEmail: 5,
  editDate: 6,
  del: [9, 10, 11, 12, 13],
  firstDraft: 14
};

// 設定シートの情報を取得
const settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
const settingData = settings.getDataRange().getValues();
const LAST_COL = settingData[0].length;
const LAST_ROW = settingData.length;
const MESSAGE_OPTIONS = settingData.slice(2, 5).map(item => item.slice(1));
const DEL_BEFORE = MESSAGE_OPTIONS[0].slice(0, 2);
const DEL_AFTER = MESSAGE_OPTIONS[0].slice(2);
const WEBHOOK_URL = settingData[8][1];
const SUPERVISOR_ADRESS = settingData[12].slice(2, COL_ROLE.del.length + 2);
const FOLDER_URL = settingData[9][1];
const SUPERVISOR_OPTIONS = settingData.slice(13, LAST_ROW).map(item => item.slice(1, COL_ROLE.del.length + 2));
const TEAM_OPTIONS = SUPERVISOR_OPTIONS.map(item => item[0]);
