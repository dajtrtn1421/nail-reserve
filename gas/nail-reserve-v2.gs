/**
 * Nail Liberty 予約フォーム用・新規スプレッドシート＋新規 Apps Script 向けひな形
 *
 * 使い方の全体手順は同じフォルダの NEW-SHEET-AND-SCRIPT.txt を参照。
 *
 * ━ 設定（必ず確認）━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * ・SECRET_TOKEN … index.html の SECRET_TOKEN と同じ文字列にする
 * ・SPREADSHEET_ID … このスクリプトをスプレッドシートに「コンテナとして紐づけ」した場合は '' のまま。
 *   スタンドアロンで作成した場合のみ、対象ブックの ID を入れる
 * ・LINE_ACCESS_TOKEN … プロジェクトの「スクリプトのプロパティ」に設定（後述）
 * ・旧プロジェクトから doGet（空き日 JSON）をコピー … このファイル末尾の案内参照
 */

var SECRET_TOKEN = 'nail-reserve-2026';
var SHEET_NAME = 'yoyaku';
/** コンテナバインドなら空文字。スタンドアロンのみ ID を入れる */
var SPREADSHEET_ID = '';

function getTargetSpreadsheet_() {
  if (SPREADSHEET_ID && String(SPREADSHEET_ID).length > 0) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ─── LIFF からの予約追記（doPost）────────────────────────────

function doPost(e) {
  var out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      out.setContent(JSON.stringify({ error: 'no body' }));
      return out;
    }

    var data = JSON.parse(e.postData.contents);
    if (data.token !== SECRET_TOKEN) {
      out.setContent(JSON.stringify({ error: 'unauthorized' }));
      return out;
    }

    if (data.action === 'appendReservation') {
      appendReservationRow_(data);
      try {
        sendLineNotification();
      } catch (notifyErr) {
        console.error('sendLineNotification failed: ' + notifyErr);
      }
      out.setContent(JSON.stringify({ ok: true }));
      return out;
    }

    out.setContent(JSON.stringify({ error: 'unknown action' }));
    return out;
  } catch (err) {
    out.setContent(JSON.stringify({ error: String(err) }));
    return out;
  }
}

function appendReservationRow_(data) {
  var sheet = getTargetSpreadsheet_().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('sheet not found: ' + SHEET_NAME);
  }

  var sentAtCell = data.sentAtIso ? new Date(data.sentAtIso) : new Date();
  var menu = String(data.menu || '');

  sheet.appendRow([
    sentAtCell,
    data.name || '',
    data.lineDisplayName || '',
    data.date1 || '',
    data.time1 || '',
    data.date2 || '',
    data.time2 || '',
    menu,
    '',
    '',
    '',
    '',
    '',
    '',
    data.note || ''
  ]);
}

// ─── 店舗用 LINE 通知（既存ロジック）────────────────────────

function convertShort(text) {
  if (!text) return '';

  var rules = {
    '持ち込みデザインコース/ 事前画像送付・要相談  ¥8,500〜': '持ち込みデザインコース ¥8,500〜',
    'スペシャルコース': 'スペシャルコース ¥7,500〜',
    'デザインコース': 'デザインコース ¥6,500〜',
    'シンプルコース': 'シンプルコース ¥5,500〜',
    'カラーグラデーション': 'カラーグラデーション ¥5,500〜',
    'ラメグラデーション': 'ラメグラデーション ¥5,000',
    'ワンカラー': 'ワンカラー ¥5,000',
    '指先スッキリケア/ 爪の長さ調整・甘皮・角質除去  ¥4,000': '指先スッキリケア ¥4,000',
    '当店ジェルオフのみ': '当店ジェルオフのみ ¥2,000',
    '親指デザインコース/ 親指デザイン+ワンカラー  ¥6,000〜': '親指デザインコース ¥6,000〜',
    '指先スッキリケア/ 爪の長さ調整・甘皮・角質除去  ¥3,000': '指先スッキリケア ¥3,000',
    '足裏つるんケア': '足裏つるんケア ¥3,500',
    '足元フルケアセット': '足元フルケアセット ¥6,000',
    '巻き爪ジェル(1本)/ 爪の湾曲を緩和  ¥6,500〜': '巻き爪ジェル(1本) ¥6,500〜',
    '巻き爪ジェル(追加1本)/ 2本以上施術時に追加  ¥3,000〜': '巻き爪ジェル(追加1本) ¥3,000〜',
    '亀裂が入っている爪がある': 'あり ¥300/本',
    '亀裂が入っている爪はない': 'なし'
  };

  for (var key in rules) {
    if (Object.prototype.hasOwnProperty.call(rules, key) && String(text).indexOf(key) !== -1) {
      return rules[key];
    }
  }
  return text;
}

function formatDate(date) {
  var options = {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: 'numeric',
    minute: 'numeric',
    second: 'numeric',
    timeZone: 'Asia/Tokyo'
  };
  return new Date(date).toLocaleString('ja-JP', options)
    .replace('年', '年 ')
    .replace('月', '月')
    .replace('日', '日')
    .replace('時', '時')
    .replace('分', '分')
    .replace('秒', '秒');
}

/**
 * 未送信の行を店舗用 LINE に送る（Messaging API broadcast）。
 * スクリプトのプロパティ: LINE_ACCESS_TOKEN（旧プロジェクトと同じトークンで同じ公式アカウントに届く）
 * LAST_SENT_ROW で二重送信を防ぐ
 */
function sendLineNotification() {
  var ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');
  if (!ACCESS_TOKEN) {
    throw new Error('LINE_ACCESS_TOKEN がスクリプトのプロパティに未設定です');
  }

  var sheet = getTargetSpreadsheet_().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();

  var lastSentRow = PropertiesService.getScriptProperties().getProperty('LAST_SENT_ROW');
  var startRow = lastSentRow ? parseInt(lastSentRow, 10) + 1 : 2;

  if (lastRow >= startRow) {
    var messages = [];
    for (var i = startRow; i <= lastRow; i++) {
      var aValue = sheet.getRange(i, 1).getValue();
      var bValue = sheet.getRange(i, 2).getValue();
      var cValue = sheet.getRange(i, 3).getValue();
      var dValue = sheet.getRange(i, 4).getValue();
      var eValue = sheet.getRange(i, 5).getValue();
      var fValue = sheet.getRange(i, 6).getValue();
      var gValue = sheet.getRange(i, 7).getValue();
      var hValue = sheet.getRange(i, 8).getValue();
      var iValue = sheet.getRange(i, 9).getValue();
      var jValue = sheet.getRange(i, 10).getValue();
      var kValue = sheet.getRange(i, 11).getValue();
      var lValue = sheet.getRange(i, 12).getValue();
      var mValue = sheet.getRange(i, 13).getValue();
      var nValue = sheet.getRange(i, 14).getValue();
      var oValue = sheet.getRange(i, 15).getValue();

      if (aValue) {
        var formattedDate = formatDate(aValue);
        messages.push(
          '予約依頼が届きました。\n\nお名前：' + bValue + '\n\nLINE名：' + cValue +
            '\n\n第１希望日：' + dValue + '\n時間：' + eValue + '\n\n第２希望日：' + fValue +
            '\n時間：' + gValue + '\n\nメニュー：' + hValue +
            '\n\nその他：' + oValue + '\n\n' + formattedDate + 'に予約がありました。'
        );
      }
    }

    if (messages.length > 0) {
      var message = messages.join('\n\n');
      var options = {
        method: 'post',
        headers: {
          Authorization: 'Bearer ' + ACCESS_TOKEN,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          messages: [{ type: 'text', text: message }]
        })
      };

      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/broadcast', options);
      PropertiesService.getScriptProperties().setProperty('LAST_SENT_ROW', String(lastRow));
    }
  }
}

/*
 * ━ doGet（カレンダー空き日 JSON）━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * 旧 Apps Script プロジェクトから、空き日を返している doGet(e) 関数を丸ごとコピーし、
 * このプロジェクトの Code.gs（または別 .gs）に貼り付けてください。
 * 関数名 doGet はプロジェクト内で 1 つだけにしてください。
 */
