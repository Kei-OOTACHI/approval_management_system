//この記事の完全コピペ======>https://style.biglobe.co.jp/entry/2023/08/30/100000
/**
 * Google APIを呼び出します。
 *
 * @param {string} url - APIのエンドポイント。
 * @param {string} oauthToken - OAuthトークン。
 * @param {string} [method='get'] - HTTPメソッド（getまたはpost）。
 * @param {Object} [payload=undefined] - 送信するデータ。
 * @throws ステータスコードが200でない場合、またはJSON形式に変換できない場合にエラーをスローします。
 * @returns {Object} APIのレスポンス。
 */
function callGoogleApi_(url, oauthToken, method = 'get', payload = undefined) {
  console.log(`callGoogleApi url=${url}, method=${method}, payload=${JSON.stringify(payload)}`);
  const options = {
    method,
    headers: { 'Authorization': `Bearer ${oauthToken}` },
    contentType: 'application/json',
  }
  if (payload) options.payload = JSON.stringify(payload);

  options['muteHttpExceptions'] = true;
  const res = UrlFetchApp.fetch(url, options);
  const responseCode = res.getResponseCode();
  const contentText = res.getContentText();
  if (responseCode !== 200) {
    const error = new Error(contentText);
    error.code = responseCode;
    if (error.code == 400) setDialog_("通知送信失敗", "自分に自分でDMすることはできないよ!");
    throw error;
  }

  let json;
  try {
    json = JSON.parse(contentText);
  } catch (e) {
    throw new Error(`JSON形式に変換できませんでした: ${contentText}`);
  }

  console.log(`response=${JSON.stringify(json, null, 2)}`);
  return json;
}

/**
 * ユーザー認証を使用してGoogle APIを呼び出します。
 *
 * @param {string} url - APIのエンドポイント。
 * @param {string} method - HTTPメソッド（getまたはpost）。
 * @param {Object} payload - 送信するデータ。
 * @returns {Object} APIのレスポンス。
 */
function callGoogleApiWithUserAuth_(url, method, payload) {
  return callGoogleApi_(url, ScriptApp.getOAuthToken(), method, payload);
}

/**
 * 指定された相手とのダイレクトメッセージのスペースを設定します（あれば既存のスペースを、なければ新しいスペースを返却）。
 * @param {string} emailOrId - 相手のメールアドレスもしくはID
 * @returns {Object} APIのレスポンス（スペースの情報）。
 */
function setupDirectMessage_(emailOrId) {
  const url = 'https://chat.googleapis.com/v1/spaces:setup';
  const payload = {
    space: {
      spaceType: 'DIRECT_MESSAGE',
      singleUserBotDm: false
    },
    memberships: [{
      member: {
        name: 'users/' + emailOrId,
        type: 'HUMAN'
      }
    }]
  };
  try {
    return callGoogleApiWithUserAuth_(url, 'post', payload);
  } catch (error) {
    if (error.code === 500) {
      setDialog_('通知送信失敗', '自分自身にはダイレクトメッセージを送信できません。')
      throw new Error(`自分自身にはダイレクトメッセージを送信できません。${error.message}`);
    }
    throw error;
  }
}

/**
 * 指定したスレッド宛にメッセージを作成します。
 *
 * @param {string} text - メッセージのテキスト。
 * @param {string} spaceName - スペースの名前。
 * @param {string} [threadName] - スレッドの名前。
 * @returns {Object} APIのレスポンス（送信結果）。
 */
function createMessageWithUserAuth_(text, spaceName, threadName) {
  let url = 'https://chat.googleapis.com/v1/' + spaceName + '/messages';
  const payload = { text };
  if (threadName) {
    url += '?messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD';
    payload.thread = {
      name: threadName
    }
  }
  return callGoogleApiWithUserAuth_(url, 'post', payload);
}

/**
 * ダイレクトメッセージを送信する
 * 
 * @param {string} emailOrId - 相手のメールアドレスもしくはID
 * @param {string} text - メッセージのテキスト。
 */
function sendDirectMessage(emailOrId, text) {
  Logger.log(emailOrId);
  const space = setupDirectMessage_(emailOrId);
  createMessageWithUserAuth_(text, space.name);
}

function testSendDirectMessage() {
  const email = 'somu@wasedasai.net';
  sendDirectMessage(email, 'こんにちは。テストです。');
}