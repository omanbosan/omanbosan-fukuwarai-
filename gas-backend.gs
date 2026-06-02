// ================================================================
// おまんぼさんイラストゲーム バックエンド
// Google Apps Script Web App
//
// 【セットアップ手順】
// 1. Google スプレッドシートを新規作成し、そのIDをSHEET_IDに貼る
// 2. Google Drive で画像保存用フォルダを作成し、そのIDをDRIVE_FOLDER_IDに貼る
// 3. 受信したいメールアドレスをOWNER_EMAILに入力
// 4. このファイルを Google Apps Script に貼り付け（script.google.com）
// 5. 「デプロイ」→「新しいデプロイ」→種類:ウェブアプリ
//    - 実行ユーザー: 自分
//    - アクセスできるユーザー: 全員
// 6. デプロイ後に表示されるURLをgame.htmlのWEBHOOK_URLに貼る
// ================================================================

const CONFIG = {
  sheetId:       'YOUR_GOOGLE_SHEET_ID',      // SpreadsheetのID（URLのd/〜/editの部分）
  ownerEmail:    'YOUR_EMAIL@gmail.com',       // 通知メール受け取り先
  driveFolderId: 'YOUR_GOOGLE_DRIVE_FOLDER_ID' // 画像保存先フォルダID
};

// GET: ランキング取得
function doGet(e) {
  try {
    const type = e && e.parameter && e.parameter.type;
    let result;
    if (type === 'ranking') {
      result = getRanking();
    } else {
      result = { ok: true, message: 'おまんぼさんイラストゲームAPI' };
    }
    return buildResponse(result);
  } catch(err) {
    return buildResponse({ error: err.message });
  }
}

// POST: スコア登録 / お絵描き投稿
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    if (data.type === 'score')   result = saveScore(data);
    else if (data.type === 'drawing') result = saveDrawing(data);
    else result = { ok: true };
    return buildResponse(result);
  } catch(err) {
    return buildResponse({ error: err.message });
  }
}

function buildResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ランキング上位20件を返す
function getRanking() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet   = ss.getSheetByName('scores');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return rows
    .map(r => ({ name: r[0], score: Number(r[1]), date: r[2], instagram: r[3] || '' }))
    .filter(r => r.name && r.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 20);
}

// スコアをスプレッドシートに保存
function saveScore(data) {
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('scores');
  if (!sheet) {
    sheet = ss.insertSheet('scores');
    sheet.appendRow(['名前', 'スコア', '日付', 'Instagram']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    (data.name || '名無し').slice(0, 20),
    Number(data.score) || 0,
    new Date().toLocaleDateString('ja-JP'),
    (data.instagram || '').replace('@', '')
  ]);
  return { ok: true };
}

// お絵描きをDriveに保存してメール通知
function saveDrawing(data) {
  if (!data.image) return { error: 'no image data' };
  const ig       = (data.instagram || 'anonymous').replace('@', '');
  const base64   = data.image.replace(/^data:image\/\w+;base64,/, '');
  const filename = `@${ig}_${Date.now()}.png`;

  // Google Driveに保存
  const folder = DriveApp.getFolderById(CONFIG.driveFolderId);
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', filename);
  const file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileUrl = file.getUrl();

  // オーナーにメール通知（画像添付）
  MailApp.sendEmail({
    to: CONFIG.ownerEmail,
    subject: `【お絵描き新着】@${ig} さんから投稿が届きました！`,
    htmlBody: `
      <h2 style="color:#27ae60;">🎨 新しいイラストが届きました！</h2>
      <table style="border-collapse:collapse;">
        <tr><td style="padding:4px 12px;font-weight:bold;">Instagram:</td><td>@${ig}</td></tr>
        <tr><td style="padding:4px 12px;font-weight:bold;">投稿日時:</td><td>${new Date().toLocaleString('ja-JP')}</td></tr>
      </table>
      <p><a href="${fileUrl}" style="background:#3498db;color:white;padding:8px 16px;border-radius:6px;text-decoration:none;">▶ 画像を確認する（Googleドライブ）</a></p>
    `,
    attachments: [blob]
  });

  // お絵描き投稿シートに記録
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let dSheet = ss.getSheetByName('drawings');
  if (!dSheet) {
    dSheet = ss.insertSheet('drawings');
    dSheet.appendRow(['Instagram', '投稿日時', 'ファイルURL']);
    dSheet.setFrozenRows(1);
  }
  dSheet.appendRow([`@${ig}`, new Date().toLocaleString('ja-JP'), fileUrl]);

  return { ok: true, fileUrl };
}
