/**
 * @OnlyCurrentDoc
 *
 * Gemini APIを利用した車検証の承認ワークフロー（Webアプリ版）
 * 1. 新規ファイルをAIで解析し、スプレッドシート（台帳）に記録
 * 2. Googleチャットに承認用WebアプリのURLを通知
 * 3. Webアプリ上で承認・修正
 * 4. ファイルをリネームし、Chatworkに通知
 */

// ▼▼▼ 設定項目 ▼▼▼
// 1. AIによる解析を行いたいフォルダのID
const ai_AI_PROCESSING_FOLDER_ID = "1wQNGTkttmTfc1dCDSORNT5B34JYxMquf";
// 2. 承認作業の台帳（スプレッドシート）のID
const ai_SPREADSHEET_ID = "1CsBH_ZRhID7ahWMdizVaPN-gfBpUrVtQF409dvR2vgw";
// 3. スプレッドシートのシート名
const ai_SHEET_NAME = "シート1";
// 5. 承認リネーム後のファイル移動先（自動リネームフォルダ）のID
const ai_AUTO_RENAME_FOLDER_ID = "XXXXXXXXXX_自動リネームフォルダIDをここに設定";

// 4. スクリプトプロパティ（プロジェクトの設定）に以下を保存
//    - GEMINI_API_KEY
//    - CHATWORK_API_TOKEN
//    - CHATWORK_ROOM_ID
//    - GOOGLE_CHAT_SPACE_ID
// ▲▲▲ 設定項目 ▲▲▲

// Gemini APIのエンドポイントURL
const ai_GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=";


// --- Webアプリのメインエントリポイント ---

/**
 * WebアプリのURLにアクセスしたときに呼び出される関数 (GETリクエスト)
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlOutput} 承認ページのHTML
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('approval_page')
    .setTitle('車検証ファイル承認')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- サーバーサイド関数 (Webアプリから呼び出される) ---

/**
 * Webアプリ（approval_page.html）から「承認待ち」のデータを取得するために呼び出される関数
 * @return {Array} 承認待ちのアイテムの配列
 */
function ai_getApprovalItems() {
  let data;
  try {
    const lastRowResponse = Sheets.Spreadsheets.Values.get(ai_SPREADSHEET_ID, `${ai_SHEET_NAME}!F:F`);
    const lastRowIndex = lastRowResponse.values ? lastRowResponse.values.length : 1;

    if (lastRowIndex < 2) {
      return []; // データ行が存在しない
    }

    const range = `${ai_SHEET_NAME}!A2:F${lastRowIndex}`;
    const response = Sheets.Spreadsheets.Values.get(ai_SPREADSHEET_ID, range);
    data = response.values;
    if (!data) {
      return []; // データが空の場合
    }
  } catch (e) {
    if (e.message.includes("was not found") || e.message.includes("empty")) {
      console.log("スプレッドシートは空か、シート名が正しくありません。");
      return [];
    }
    console.error(`Sheets API (get items) 実行エラー: ${e}`);
    return []; // エラー時は空の配列
  }

  const items = [];
  for (let i = 0; i < data.length; i++) {
    // data[i]は配列 [タイムスタンプ, 元のファイル名, AIによる推奨名, ファイルID, ファイルURL, ステータス]
    if (data[i][5] === "承認待ち") {
      items.push({
        row: i + 2, // スプレッドシートの実際の行番号
        originalName: data[i][1],
        suggestedName: data[i][2],
        fileUrl: data[i][4]
      });
    }
  }
  return items;
}

/**
 * Webアプリ（approval_page.html）から承認ボタンが押されたときに呼び出される関数
 * @param {Object} approvalData - 承認データ ({ row, newName })
 * @return {String} 処理結果
 */
function ai_processApproval(approvalData) {
  const { row, newName } = approvalData;
  
  try {
    // スプレッドシートから該当行のデータを取得（Sheets APIを使用）
    const getRange = `${ai_SHEET_NAME}!D${row}`; // ファイルID（D列）のみ取得
    const dataRow = Sheets.Spreadsheets.Values.get(ai_SPREADSHEET_ID, getRange).values;
    if (!dataRow || !dataRow[0]) {
      throw new Error("スプレッドシートから対象の行データが見つかりません。");
    }
    const fileId = dataRow[0][0];
    
    // Drive APIでファイル操作
    const file = DriveApp.getFileById(fileId);
    
    let originalExtension = '';
    const lastDotIndex = file.getName().lastIndexOf('.');
    if (lastDotIndex > 0 && lastDotIndex < file.getName().length - 1) {
      originalExtension = file.getName().substring(lastDotIndex);
    }
    
    const finalNewName = `${newName}[R]${originalExtension}`;
    file.setName(finalNewName);
    
    // スプレッドシートを更新（Sheets APIを使用）
    const updateRange = `${ai_SHEET_NAME}!C${row}:F${row}`; // C列（推奨名）からF列（ステータス）を更新
    const values = [[
      newName, // C列: 修正された可能性のある名前で更新
      fileId,  // D列: (変更なし)
      file.getUrl(), // E列: (変更なし)
      "処理完了" // F列: ステータス
    ]];
    Sheets.Spreadsheets.Values.update({ values: values }, ai_SPREADSHEET_ID, updateRange, { valueInputOption: "RAW" });
    
    console.log(`リネーム成功: ${fileId} -> ${finalNewName}`);

    // 自動リネームフォルダへ移動（自動版との出力先統一）
    const destFolder = DriveApp.getFolderById(ai_AUTO_RENAME_FOLDER_ID);
    file.moveTo(destFolder);

    // Chatworkに通知
    ai_postToChatwork(file);

    return "承認・リネームが完了しました。";

  } catch (e) {
    // エラー発生時にステータスを更新
    const updateRange = `${ai_SHEET_NAME}!F${row}`; // F列（ステータス）のみ更新
    const values = [["エラー発生"]];
    Sheets.Spreadsheets.Values.update({ values: values }, ai_SPREADSHEET_ID, updateRange, { valueInputOption: "RAW" });
    
    console.error(`リネーム失敗: ${e.toString()}`);
    return `エラーが発生しました: ${e.message}`;
  }
}


// --- 定期実行される関数 ---

/**
 * フォルダを監視し、新しいファイルをスプレッドシートとGoogleチャットに通知する関数
 */
function ai_processNewFilesForApproval() {
  const folder = DriveApp.getFolderById(ai_AI_PROCESSING_FOLDER_ID);
  
  // スプレッドシートから既存のファイルIDリストを取得（Sheets APIを使用）
  let existingFileIds = [];
  try {
    const lastRowResponse = Sheets.Spreadsheets.Values.get(ai_SPREADSHEET_ID, `${ai_SHEET_NAME}!F:F`);
    const lastRowIndex = lastRowResponse.values ? lastRowResponse.values.length : 1; // ヘッダー行のみの場合は1

    if (lastRowIndex >= 2) {
      const range = `${ai_SHEET_NAME}!D2:D${lastRowIndex}`; // D列（ファイルID）を取得
      const response = Sheets.Spreadsheets.Values.get(ai_SPREADSHEET_ID, range);
      if (response.values) {
        existingFileIds = response.values.flat().filter(String);
      }
    }
  } catch (e) {
    if (e.message.includes("was not found") || e.message.includes("empty")) {
      console.log("スプレッドシートは空か、シート名が正しくありません。処理を続行します。");
      existingFileIds = [];
    } else {
      console.error(`Sheets API (get file IDs) 実行エラー: ${e}`);
      return; // エラー時は実行を中断
    }
  }

  const files = folder.getFiles();
  while(files.hasNext()){
    const file = files.next();
    const fileId = file.getId();

    // スプレッドシートに存在しないファイルのみ処理
    if(!existingFileIds.includes(fileId)){
      console.log(`新しいファイルを発見: ${file.getName()}`);
      
      const suggestedName = ai_getSuggestedNameFromGemini(file);
      
      let newRow;
      if (suggestedName.startsWith("エラー")) {
        // AI解析失敗
        newRow = [
          new Date(), file.getName(), suggestedName,
          fileId, file.getUrl(), "AIエラー"
        ];
        ai_postToGoogleChat(`[AI解析エラー] ファイル: ${file.getName()}`, null); // テキストのみ通知
      } else {
        // AI解析成功
        newRow = [
          new Date(), file.getName(), suggestedName,
          fileId, file.getUrl(), "承認待ち"
        ];
        
        // ユーザーが確認した「アクセス可能なURL」を直接指定します。
        const webAppUrl = "https://script.google.com/a/macros/h-rent.com/s/AKfycbxdZpazcckSsf4-zkwHS-pNVdhQgCN71FdehabWTWIYC-u6kRfLVmwyyQ9CIHp5hwiZgQ/exec";
        ai_postToGoogleChat(`[承認待ち] ${file.getName()} のファイル名を確認してください。`, webAppUrl);
      }
      
      // スプレッドシートに行を追加（Sheets APIを使用）
      try {
        const range = `${ai_SHEET_NAME}!A:F`;
        Sheets.Spreadsheets.Values.append({ values: [newRow] }, ai_SPREADSHEET_ID, range, { valueInputOption: "USER_ENTERED" });
      } catch (e) {
        console.error(`Sheets API (append row) 実行エラー: ${e}`);
      }
    }
  }
}

/**
 * Gemini APIに画像を送信し、推奨ファイル名を取得する関数。
 */
function ai_getSuggestedNameFromGemini(file) {
  try {
    // ★★ 更新: 新しいプロンプトをここに反映 ★★
    const prompt = `
      この自動車検査証の画像から以下の情報を抽出してください。
      1.  **記録年月日**: (例: 令和6年10月24日)
      2.  **交付年月日**: (例: 令和6年10月20日)
      3.  **使用者の氏名又は名称**:
      4.  **所有者の氏名又は名称**:
      5.  **登録番号**: (例: 品川 300 わ 1234)
      6.  **車台番号**:
      7.  **返納証明書フラグ**: (画像内に「返納証明書」「自動車検査証返納証明書」といった記載があれば true, なければ false)

      そして、以下のルールに厳密に従って、ファイル名を**1行の文字列**として生成してください。

      # ファイル名形式
      (日付)_(フラグ)_(使用者名)_(登録番号)_(車台番号)

      # ルール詳細

      1.  **(日付) の決定:**
          * **抹消の場合:** (ルール7の「返納証明書フラグ」が true の場合)、抽出した「交付年月日」をYYYYMMDD形式（西暦）に変換して使用します。
          * **上記以外の場合:** 抽出した「記録年月日」をYYYYMMDD形式（西暦）に変換して使用します。

      2.  **(フラグ) の決定 (オプション):**
          * **抹消の場合:** (ルール7の「返納証明書フラグ」が true の場合)、日付の直後に \`_抹消\` を追加します。
          * **更新の場合:** (抹消ではない AND 「記録年月日」と「交付年月日」が異なる日付の場合)、日付の直後に \`_更新\` を追加します。
          * **上記以外の場合:** フラグは追加しません (日付の直後は \`_\` になります)。

      3.  **(使用者名) の決定:**
          * **もし「使用者の氏名又は名称」が「***」の場合:** 「所有者の氏名又は名称」を使用します。
          * **もし「使用者の氏名又は名称」が「株式会社橋本商会」 AND 「登録番号」に「わ」が含まれる場合:** \`(株)橋本ﾚﾝﾀｶｰ\` という文字列を使用します。
          * **上記以外の場合:** 「使用者の氏名又は名称」を使用します。

      4.  **短縮・整形ルール:**
          * (使用者名)、(登録番号)、(車台番号) に含まれる全ての空白（半角・全角）は、**完全に削除**してください。
          * (使用者名) に含まれる法人格は、以下のように短縮してください。
              * \`株式会社\` → \`(株)\`
              * \`合同会社\` → \`(同)\`
              * （その他、有限会社→(有) など、一般的な法人格も同様に短縮してください）

      5.  **結合ルール:**
          * 各要素（日付、(使用者名)、(登録番号)、(車台番号)）は、必ずアンダースコア(_)で区切ってください。
          * (フラグ) が存在する場合、\`日付_フラグ_使用者名...\` のようになります。
          * (フラグ) が存在しない場合、\`日付_使用者名...\` のようになります。

      6.  **最終出力ルール:**
          * 説明、前置き、箇条書き、追加のテキストは一切含めないでください。
          * **応答は、生成されたファイル名文字列のみ**にしてください。
    `;
    
    const imageBlob = file.getBlob();
    const base64ImageData = Utilities.base64Encode(imageBlob.getBytes());
    const payload = { contents: [{ parts: [ { text: prompt }, { inlineData: { mimeType: imageBlob.getContentType(), data: base64ImageData } } ] }] };
    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return "エラー: APIキー未設定";

    const response = UrlFetchApp.fetch(`${ai_GEMINI_API_URL}${apiKey}`, options);
    const result = JSON.parse(response.getContentText());

    if (result.candidates && result.candidates[0].content.parts[0].text) {
      return result.candidates[0].content.parts[0].text.trim();
    } else {
      console.error("Gemini APIからの応答が無効です:", result);
      return "エラー: AIによる解析失敗";
    }
  } catch (e) {
    console.error(`Gemini API呼び出し中にエラー: ${e.toString()}`);
    return `エラー: ${e.message}`;
  }
}


// --- 通知関数 ---

/**
 * Googleチャットにメッセージを投稿する
 * @param {string} message - 送信するテキストメッセージ
 * @param {string | null} webAppUrl - 承認用WebアプリのURL（あれば）
 */
function ai_postToGoogleChat(message, webAppUrl) {
  const properties = PropertiesService.getScriptProperties();
  const spaceId = properties.getProperty('GOOGLE_CHAT_SPACE_ID');
  if (!spaceId) {
    console.error("Google ChatのスペースIDが設定されていません。");
    return;
  }
  
  const url = `https://chat.googleapis.com/v1/${spaceId}/messages`;
  
  let messageBody = message;
  if (webAppUrl) {
    messageBody += `\n\n▼ 承認画面を開く\n${webAppUrl}`;
  }
  
  const payload = { "text": messageBody };

  const options = {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      console.error(`Google Chatへの投稿に失敗しました: ${response.getContentText()}`);
    }
  } catch (e) {
    console.error(`Google Chat API呼び出し中にエラー: ${e.toString()}`);
  }
}


/**
 * Chatworkに完了通知を投稿する
 * @param {GoogleAppsScript.Drive.File} file - 投稿するファイルオブジェクト
 */
function ai_postToChatwork(file) { // ★★ 修正 ★★ folder 引数を削除（利用しないため）
  const properties = PropertiesService.getScriptProperties();
  const apiToken = properties.getProperty('CHATWORK_API_TOKEN');
  const roomId = properties.getProperty('CHATWORK_ROOM_ID');

  if (!apiToken || !roomId) {
    console.error("ChatworkのAPIトークンまたはルームIDが設定されていません。");
    return;
  }

  const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const fileUrl = file.getUrl();
  
  // ★★ 修正 ★★
  // フォルダリンクを削除し、ファイルリンクのみのシンプルなメッセージに変更
  const messageBody = `[info][title]承認された車検証ファイルがリネームされました[/title]${file.getName()}\n${fileUrl}[/info]`;
  
  const payload = { body: messageBody };
  const options = { method: 'post', headers: { 'X-ChatWorkToken': apiToken }, payload: payload, muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      console.log("Chatworkへのメッセージ投稿に成功しました。");
    } else {
      console.error(`Chatworkへの投稿に失敗しました: ${response.getResponseCode()}, 応答: ${response.getContentText()}`);
    }
  } catch (e) {
    console.error(`Chatwork API呼び出し中にエラー: ${e.toString()}`);
  }
}


// --- トリガー設定 ---

/**
 * 時間ベースのトリガーを作成する関数。
 */
function ai_createTimeDrivenTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === "ai_processNewFilesForApproval") {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger("ai_processNewFilesForApproval")
    .timeBased()
    .everyMinutes(10)
    .create();
  console.log("10分ごとのAI処理・Chat通知トリガーを設定しました。");
}