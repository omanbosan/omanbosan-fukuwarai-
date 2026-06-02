// ================================================================
// おまんぼさんイラストゲーム バックエンド
// Google Apps Script Web App
//
// 【セットアップ手順】
// 1. このファイルを script.google.com に貼り付け
// 2. ownerEmail に通知先メールアドレスを入力
// 3. 「デプロイ」→「新しいデプロイ」→ 種類: ウェブアプリ
//    - 実行ユーザー: 自分
//    - アクセスできるユーザー: 全員
// 4. デプロイ後のURLをgame.htmlのWEBHOOK_URLに貼る
// 5. 日次まとめメールのトリガーを設定:
//    「トリガー」→「トリガーを追加」→ 関数: sendDailySummary
//    イベントソース: 時間主導型 / 種類: 日タイマー / 時刻: 好きな時間
// ================================================================

const CONFIG = {
  sheetId:       '1nkGoMoNKc4opnD-Z_Z41I_4re8Brz7lFyBdNXJfpyIw',
  ownerEmail:    'YOUR_EMAIL@gmail.com',  // ← ここにメールアドレスを入力
  driveFolderId: '1gXwYsxpKBqZ6tMKdNyOPGVvtQTcV1QgN'
};

// ----------------------------------------------------------------
// GET: ランキング取得
// ----------------------------------------------------------------
function doGet(e) {
  try {
    const type = e && e.parameter && e.parameter.type;
    if (type === 'ranking') return buildResponse(getRanking());
    return buildResponse({ ok: true, message: 'おまんぼさんイラストゲームAPI' });
  } catch(err) {
    return buildResponse({ error: err.message });
  }
}

// ----------------------------------------------------------------
// POST: スコア登録 / お絵描き投稿
// ----------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'score')   return buildResponse(saveScore(data));
    if (data.type === 'drawing') return buildResponse(saveDrawing(data));
    return buildResponse({ ok: true });
  } catch(err) {
    return buildResponse({ error: err.message });
  }
}

function buildResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// ランキング上位20件を返す
// ----------------------------------------------------------------
function getRanking() {
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('scores');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map(r => ({ name: r[0], score: Number(r[1]), date: r[2], instagram: r[3] || '' }))
    .filter(r => r.name && r.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 20);
}

// ----------------------------------------------------------------
// スコアをスプレッドシートに保存
// ----------------------------------------------------------------
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

// ----------------------------------------------------------------
// お絵描きをDriveに保存（メールは日次まとめで送信）
// ----------------------------------------------------------------
function saveDrawing(data) {
  if (!data.image) return { error: 'no image data' };
  const ig     = (data.instagram || 'anonymous').replace('@', '');
  const base64 = data.image.replace(/^data:image\/\w+;base64,/, '');

  // Google Drive に保存
  const folder = DriveApp.getFolderById(CONFIG.driveFolderId);
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', `@${ig}_${Date.now()}.png`);
  const file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileUrl = file.getUrl();

  // drawingsシートに「未送信」ステータスで記録
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('drawings');
  if (!sheet) {
    sheet = ss.insertSheet('drawings');
    sheet.appendRow(['Instagram', '投稿日時', 'ファイルURL', '送信状態']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([`@${ig}`, new Date().toLocaleString('ja-JP'), fileUrl, '未送信']);

  return { ok: true, fileUrl };
}

// ----------------------------------------------------------------
// 日次まとめメール（毎日1回トリガーで実行）
// ----------------------------------------------------------------
function sendDailySummary() {
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet || sheet.getLastRow() < 2) return;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

  // 「未送信」の行だけ抽出
  const unsent = rows
    .map((r, i) => ({ ig: r[0], date: r[1], url: r[2], status: r[3], row: i + 2 }))
    .filter(r => r.status === '未送信');

  if (unsent.length === 0) return; // 新着なければ何もしない

  // まとめメール本文を作成
  const today = new Date().toLocaleDateString('ja-JP');
  let html = `
    <h2 style="color:#27ae60;">🎨 お絵描き投稿 日次まとめ【${today}】</h2>
    <p>本日の新着投稿：<strong>${unsent.length}件</strong></p>
    <table style="border-collapse:collapse;width:100%;">
      <tr style="background:#27ae60;color:white;">
        <th style="padding:8px 12px;">Instagram</th>
        <th style="padding:8px 12px;">投稿日時</th>
        <th style="padding:8px 12px;">画像</th>
      </tr>
      ${unsent.map((r, i) => `
        <tr style="background:${i%2===0?'#f9f9f9':'white'};">
          <td style="padding:8px 12px;">${r.ig}</td>
          <td style="padding:8px 12px;">${r.date}</td>
          <td style="padding:8px 12px;"><a href="${r.url}">画像を見る</a></td>
        </tr>
      `).join('')}
    </table>
    <p style="margin-top:16px;">
      <a href="https://drive.google.com/drive/folders/${CONFIG.driveFolderId}"
         style="background:#3498db;color:white;padding:8px 16px;border-radius:6px;text-decoration:none;">
        ▶ Driveフォルダをまとめて確認
      </a>
    </p>
  `;

  MailApp.sendEmail({
    to: CONFIG.ownerEmail,
    subject: `【おまんぼさんイラスト】${today} の投稿まとめ（${unsent.length}件）`,
    htmlBody: html
  });

  // 送信済みに更新
  unsent.forEach(r => sheet.getRange(r.row, 4).setValue('送信済み'));
}
