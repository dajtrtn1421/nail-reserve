// index.html の SECRET_TOKEN と同じ値を GAS 側に設定してください
const SECRET_TOKEN = "";

// ============================================================
// ★ 設定欄（GAS デプロイ時に値を入れる。リポジトリにはコミットしない）
// ============================================================
// お客様へpush送信用（お店の公式LINEのトークン）
const SHOP_CHANNEL_TOKEN = "";

// オーナー通知用（予約通知アカウントのトークン）
const NOTIFY_CHANNEL_TOKEN = "";
// ============================================================

function doGet(e) {
  if (!e.parameter || e.parameter.token !== SECRET_TOKEN) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "unauthorized" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName("空き日管理");
  const rows = sheet.getDataRange().getValues();

  const result = rows.slice(1).map(row => ({
    date: Utilities.formatDate(new Date(row[0]), "Asia/Tokyo", "yyyy-MM-dd"),
    available: row[1] === "○"
  }));

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "invalid_json" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.token !== SECRET_TOKEN) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "unauthorized" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === "appendReservation") {
    try {
      const debugSheet = SpreadsheetApp.getActiveSpreadsheet()
                         .getSheetByName("空き日管理");
    debugSheet.appendRow(["DEBUG", new Date(), body.lineUserId, body.lineDisplayName]);

      const customerText =
`予約をリクエストしました💅
━━━━━━━━━━━━━━
お名前：${body.name}
メニュー：${body.menu}

第一希望：${body.date1} ${body.time1}
第二希望：${body.date2} ${body.time2}${body.note ? "\n\nご要望：" + body.note : ""}
━━━━━━━━━━━━━━
現時点で、ご予約はまだ確定していません。ご注意ください。
ご予約確定についてはご連絡が入りますので、ご対応をお願いいたします。

来店希望日の前日までにLibertyから連絡がない場合、お手数ですが直接お問い合わせください。`;

      if (body.lineUserId) {
        UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
          method: "post",
          headers: {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + SHOP_CHANNEL_TOKEN  // ★ 修正
          },
          payload: JSON.stringify({
            to: body.lineUserId,
            messages: [{ type: "text", text: customerText }]
          })
        });
      }

      // ============================================================
      // ② 予約通知アカウントにbroadcast（NOTIFY_CHANNEL_TOKEN使用）
      // ============================================================
      const notifyText =
`【新規予約リクエスト】
━━━━━━━━━━━━━━
お名前：${body.name}
LINE名：${body.lineDisplayName}
メニュー：${body.menu}

第一希望：${body.date1} ${body.time1}
第二希望：${body.date2} ${body.time2}${body.note ? "\nご要望：" + body.note : ""}
━━━━━━━━━━━━━━`;

      UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", {
        method: "post",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + NOTIFY_CHANNEL_TOKEN  // ★ 修正
        },
        payload: JSON.stringify({
          messages: [{ type: "text", text: notifyText }]
        })
      });

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: "unknown_action" }))
    .setMimeType(ContentService.MimeType.JSON);
}