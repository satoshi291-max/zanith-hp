/**
 * Zanith HP — Google Apps Script Webhook
 *
 * 【セットアップ手順】
 * 1. Googleスプレッドシートを新規作成
 *    - シート1: 「問い合わせ」に名前変更
 *    - シート2: 「先行登録」に名前変更
 * 2. 「問い合わせ」シートのA1に以下をヘッダーとして入力:
 *    受付日時 | 業種 | 規模 | エリア | 相談内容 | メール | 会社名 | ステータス | スコア | 次アクション | メモ
 * 3. 「先行登録」シートのA1に以下をヘッダーとして入力:
 *    登録日時 | メール | ステータス
 * 4. 拡張機能 → Apps Script を開き、このコードを貼り付ける
 * 5. NOTIFY_EMAIL を自分のメールアドレスに変更する
 * 6. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *    - 実行ユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 7. デプロイURLをコピーして index.html の GAS_URL に貼り付ける
 */

const NOTIFY_EMAIL = 'info@zanith.jp'; // 通知先メールアドレス

// 業種コードを日本語に変換
const INDUSTRY_MAP = {
  food_service: '飲食・給食',
  hotel:        '宿泊・ホテル',
  care:         '介護・福祉',
  food_mfg:     '食品製造',
  agriculture:  '農業・農産物加工',
  retail:       '小売・EC',
  construction: '建設・設備',
  other:        'その他'
};

const SIZE_MAP = {
  small:  '〜10名',
  medium: '11〜50名',
  large:  '51名以上'
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.type === 'contact') {
      const sheet = ss.getSheetByName('問い合わせ');
      const industry = INDUSTRY_MAP[data.industry] || data.industry || '';
      const size     = SIZE_MAP[data.company_size] || data.company_size || '';

      sheet.appendRow([
        new Date(),
        industry,
        size,
        data.region  || '',
        data.message || '',
        data.email   || '',
        data.company || '',
        '未対応',   // ステータス
        '',         // スコア（手動入力）
        '',         // 次アクション
        ''          // メモ
      ]);

      GmailApp.sendEmail(
        NOTIFY_EMAIL,
        '【Zanith】新しい無料診断依頼が届きました',
        [
          '業種: '   + industry,
          '規模: '   + size,
          'エリア: ' + (data.region  || ''),
          'メール: ' + (data.email   || ''),
          '会社名: ' + (data.company || '（未入力）'),
          '',
          '相談内容:',
          data.message || ''
        ].join('\n')
      );

    } else if (data.type === 'preregister') {
      const sheet = ss.getSheetByName('先行登録');
      sheet.appendRow([
        new Date(),
        data.email || '',
        '未連絡'
      ]);

      GmailApp.sendEmail(
        NOTIFY_EMAIL,
        '【Zanith】先行登録がありました',
        'メール: ' + (data.email || '')
      );
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
