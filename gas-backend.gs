// ================================================================
// おまんぼさんイラストゲーム バックエンド v1.6
//
// 【セットアップ手順】
// 1. このコードをすべて貼り替えて「デプロイ」→「新しいデプロイ」（毎回新しいデプロイが必要）
//    ※URLが変わる場合はgame.htmlのWEBHOOK_URLも更新
// 2. 日次まとめトリガー: sendDailySummary / 時間主導型 / 日タイマー
// 3. 月次リセットトリガー: archiveMonthlyRanking / 時間主導型 / 月タイマー / 毎月1日
// ================================================================

const CONFIG = {
  sheetId:       '1nkGoMoNKc4opnD-Z_Z41I_4re8Brz7lFyBdNXJfpyIw',
  ownerEmail:    'omanbosan.lv@gmail.com',
  driveFolderId: '1gXwYsxpKBqZ6tMKdNyOPGVvtQTcV1QgN'
};

// ----------------------------------------------------------------
// GET
// ----------------------------------------------------------------
function doGet(e) {
  try {
    const type = e && e.parameter && e.parameter.type;
    if (type === 'ranking') return buildResponse(getRanking());
    if (type === 'archive') return buildResponse(getLastMonthRanking());
    if (type === 'zukan')   return buildResponse(getZukan());
    if (type === 'pending') return buildResponse(getPending());
    return buildResponse({ ok: true, message: 'おまんぼさんイラストゲームAPI v1.7' });
  } catch(err) {
    return buildResponse({ error: err.message });
  }
}

// ----------------------------------------------------------------
// POST
// ----------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'score')      return buildResponse(saveScore(data));
    if (data.type === 'drawing')    return buildResponse(saveDrawing(data));
    if (data.type === 'setApproval') return buildResponse(setApproval(data));
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
// 今月のランキング
// ----------------------------------------------------------------
function getRanking() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('scores');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map(r => ({ name: r[0], score: Number(r[1]), date: r[2], instagram: r[3] || '' }))
    .filter(r => r.name && r.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 20);
}

// ----------------------------------------------------------------
// 先月のランキング（アーカイブから取得）
// ----------------------------------------------------------------
function getLastMonthRanking() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('scores_archive');
  if (!sheet || sheet.getLastRow() < 2) return [];

  // 最新の月ラベルを取得
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const latestMonth = rows[rows.length - 1][0];

  return rows
    .filter(r => r[0] === latestMonth)
    .map(r => ({ month: r[0], name: r[2], score: Number(r[3]), instagram: r[4] || '' }))
    .sort((a, b) => b.score - a.score);
}

// ----------------------------------------------------------------
// スコア保存
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
// 審査中一覧取得（管理画面用）
// ----------------------------------------------------------------
function getPending() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .map((r, i) => ({
      row:         i + 2,
      instagram:   (r[0] || '').replace('@', ''),
      date:        r[1] ? String(r[1]).slice(0, 10) : '',
      imageUrl:    driveUrlToThumb(r[2]),
      villageName: r[4] || '',
      comment:     r[5] || '',
      approved:    r[7] || '審査中'
    }))
    .filter(r => r.villageName);
}

// ----------------------------------------------------------------
// 承認・却下（管理画面用）
// ----------------------------------------------------------------
function setApproval(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet) return { error: 'no sheet' };
  const row    = Number(data.row);
  const status = data.status === '承認済み' ? '承認済み' : '却下';
  sheet.getRange(row, 8).setValue(status);
  return { ok: true, row, status };
}

// ----------------------------------------------------------------
// DriveのURLをサムネイル表示用URLに変換
// ----------------------------------------------------------------
function driveUrlToThumb(url) {
  if (!url || !url.startsWith('http')) return '';
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w300';
  return '';
}

// ----------------------------------------------------------------
// 図鑑取得（承認済みのみ）
// ----------------------------------------------------------------
function getZukan() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .map((r, i) => ({
      no:          i + 1,
      instagram:   (r[0] || '').replace('@', ''),
      date:        r[1] ? String(r[1]).slice(0, 10) : '',
      imageUrl:    driveUrlToThumb(r[2]),
      villageName: r[4] || '',
      comment:     r[5] || '',
      approved:    r[7] || ''
    }))
    .filter(r => r.villageName && r.approved === '承認済み')
    .reverse(); // 新しい順
}

// ----------------------------------------------------------------
// お絵描き保存（シート記録→Drive保存の順で確実に残す）
// ----------------------------------------------------------------
function saveDrawing(data) {
  if (!data.image) return { error: 'no image data' };
  const ig          = String(data.instagram   || 'anonymous').replace('@', '');
  const villageName = String(data.villageName || '').slice(0, 10);
  const comment     = String(data.comment     || '').slice(0, 20);

  // 1. シートに仮記録（Drive保存前でも記録が残るように）
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('drawings');
  if (!sheet) {
    sheet = ss.insertSheet('drawings');
    sheet.appendRow(['Instagram', '投稿日時', 'ファイルURL', '送信状態', '村人名', 'ひとこと', '図鑑No', '承認']);
    sheet.setFrozenRows(1);
  }
  // ヘッダー列が足りない場合は拡張
  const colCount = sheet.getLastColumn();
  if (colCount < 5) sheet.getRange(1, 5).setValue('村人名');
  if (colCount < 6) sheet.getRange(1, 6).setValue('ひとこと');
  if (colCount < 7) sheet.getRange(1, 7).setValue('図鑑No');
  if (colCount < 8) sheet.getRange(1, 8).setValue('承認');

  const zukanNo  = Math.max(sheet.getLastRow(), 1);
  const rowIndex = sheet.getLastRow() + 1;
  sheet.appendRow([
    '@' + ig,
    new Date().toLocaleString('ja-JP'),
    'Drive保存中...',
    '未送信',
    villageName,
    comment,
    zukanNo,
    '審査中'  // ← デフォルトは審査中。承認済み にすると図鑑に表示
  ]);

  // 2. Drive に保存
  try {
    const base64 = data.image.replace(/^data:image\/\w+;base64,/, '');
    const folder = DriveApp.getFolderById(CONFIG.driveFolderId);
    const blob   = Utilities.newBlob(
      Utilities.base64Decode(base64), 'image/png', `@${ig}_${Date.now()}.png`
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();

    // 3. シートのURLを更新
    sheet.getRange(rowIndex, 3).setValue(fileUrl);
    return { ok: true, fileUrl };
  } catch(err) {
    // Drive保存失敗でもシートにエラーを記録
    sheet.getRange(rowIndex, 3).setValue('Drive保存失敗: ' + err.message);
    sheet.getRange(rowIndex, 4).setValue('エラー');
    return { error: err.message };
  }
}

// ----------------------------------------------------------------
// 日次まとめメール（毎日1回トリガーで実行）
// ----------------------------------------------------------------
function sendDailySummary() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet || sheet.getLastRow() < 2) return;

  const rows   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const unsent = rows
    .map((r, i) => ({ ig: r[0], date: r[1], url: r[2], status: r[3], name: r[4], comment: r[5], approved: r[7], row: i + 2 }))
    .filter(r => r.status === '未送信');

  if (unsent.length === 0) return;

  const today = new Date().toLocaleDateString('ja-JP');
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.sheetId}/edit`;
  MailApp.sendEmail({
    to: CONFIG.ownerEmail,
    subject: `【おまんぼさんイラスト】${today} の投稿まとめ（${unsent.length}件）`,
    htmlBody: `
      <h2 style="color:#27ae60;">🎨 お絵描き投稿 日次まとめ【${today}】</h2>
      <p>本日の新着：<strong>${unsent.length}件</strong></p>
      <p style="background:#fff3cd;padding:10px;border-radius:6px;font-size:0.9em;">
        📋 <b>図鑑に載せる場合は「承認」列を <span style="color:green">承認済み</span> に変更してください</b><br>
        問題がある絵は <span style="color:red">却下</span> と入力すると図鑑に表示されません。
      </p>
      <table style="border-collapse:collapse;width:100%;">
        <tr style="background:#27ae60;color:white;">
          <th style="padding:8px;">村人名</th>
          <th style="padding:8px;">ひとこと</th>
          <th style="padding:8px;">Instagram</th>
          <th style="padding:8px;">画像</th>
          <th style="padding:8px;">承認状態</th>
        </tr>
        ${unsent.map((r, i) => `
          <tr style="background:${i%2===0?'#f9f9f9':'white'};">
            <td style="padding:8px;font-weight:bold;">${r.name || '（未入力）'}</td>
            <td style="padding:8px;color:#666;">${r.comment ? '「'+r.comment+'」' : '-'}</td>
            <td style="padding:8px;">${r.ig}</td>
            <td style="padding:8px;"><a href="${r.url}" style="color:#3498db;">画像を見る🔍</a></td>
            <td style="padding:8px;color:#e67e22;font-weight:bold;">${r.approved || '審査中'}</td>
          </tr>
        `).join('')}
      </table>
      <div style="margin-top:16px;display:flex;gap:10px;flex-wrap:wrap;">
        <a href="${sheetUrl}" style="background:#27ae60;color:white;padding:10px 18px;border-radius:6px;text-decoration:none;display:inline-block;">
          📊 スプレッドシートで承認する
        </a>
        <a href="https://drive.google.com/drive/folders/${CONFIG.driveFolderId}" style="background:#3498db;color:white;padding:10px 18px;border-radius:6px;text-decoration:none;display:inline-block;">
          📂 Driveフォルダを確認
        </a>
      </div>
    `
  });

  unsent.forEach(r => sheet.getRange(r.row, 4).setValue('送信済み'));
}

// ----------------------------------------------------------------
// 月次ランキングアーカイブ（毎月1日トリガーで実行）
// ----------------------------------------------------------------
function archiveMonthlyRanking() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('scores');
  if (!sheet || sheet.getLastRow() < 2) return;

  const lastMonth = new Date();
  lastMonth.setMonth(lastMonth.getMonth() - 1);
  const label = `${lastMonth.getFullYear()}年${lastMonth.getMonth() + 1}月`;

  // アーカイブシートに保存
  let archSheet = ss.getSheetByName('scores_archive');
  if (!archSheet) {
    archSheet = ss.insertSheet('scores_archive');
    archSheet.appendRow(['月', '順位', '名前', 'スコア', 'Instagram']);
    archSheet.setFrozenRows(1);
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map(r => ({ name: r[0], score: Number(r[1]), instagram: r[3] || '' }))
    .filter(r => r.name && r.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 10);

  rows.forEach((r, i) => {
    archSheet.appendRow([label, i + 1, r.name, r.score, r.instagram]);
  });

  // 先月結果をメールで送信
  const medals = ['🥇', '🥈', '🥉'];
  MailApp.sendEmail({
    to: CONFIG.ownerEmail,
    subject: `【${label} ランキング確定】おまんぼさんイラストゲーム`,
    htmlBody: `
      <h2>🏆 ${label} ランキング確定結果</h2>
      <table style="border-collapse:collapse;width:100%;">
        <tr style="background:#9b59b6;color:white;">
          <th style="padding:8px;">順位</th><th style="padding:8px;">名前</th>
          <th style="padding:8px;">スコア</th><th style="padding:8px;">Instagram</th>
        </tr>
        ${rows.map((r, i) => `
          <tr style="background:${i%2===0?'#f9f9f9':'white'};">
            <td style="padding:8px;text-align:center;">${medals[i] || i+1}</td>
            <td style="padding:8px;">${r.name}</td>
            <td style="padding:8px;text-align:center;">${r.score}点</td>
            <td style="padding:8px;">${r.instagram ? '@'+r.instagram : '-'}</td>
          </tr>
        `).join('')}
      </table>
      <p style="color:#888;">※ ランキングは本日リセットされました</p>
    `
  });

  // scoresシートをクリア（ヘッダー以外）
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
}

// ----------------------------------------------------------------
// 動作テスト用
// ----------------------------------------------------------------
function testDrive() {
  const folder = DriveApp.getFolderById(CONFIG.driveFolderId);
  Logger.log('フォルダ名: ' + folder.getName());
  Logger.log('✅ Drive接続OK');
}

function testMail() {
  MailApp.sendEmail({
    to: CONFIG.ownerEmail,
    subject: '【テスト】おまんぼさんイラストゲーム GAS接続確認',
    body: 'GASのメール送信テストです。このメールが届いていれば設定完了です！'
  });
  Logger.log('✅ メール送信OK');
}

// ----------------------------------------------------------------
// 【初回のみ実行】承認列をセットアップ
// drawingsシートにH列「承認」を追加し、既存データを全て「承認済み」にする
// ----------------------------------------------------------------
function setupApprovalColumn() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet) { Logger.log('❌ drawingsシートが見つかりません'); return; }

  const lastRow = sheet.getLastRow();

  // H1にヘッダー追加
  sheet.getRange(1, 8).setValue('承認');

  // 既存データ（2行目以降）をすべて「承認済み」に設定
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 8, lastRow - 1, 1);
    const values = Array(lastRow - 1).fill(['承認済み']);
    range.setValues(values);
  }

  Logger.log(`✅ 承認列セットアップ完了！${lastRow - 1}件を「承認済み」に設定しました`);
}
