/**
 * genius-pc-screen / genius-pc-remote 用バックエンド (Google Apps Script)
 *
 * セットアップ手順:
 * 1. 新しいGoogleスプレッドシートを作成する
 * 2. メニュー「拡張機能」>「Apps Script」を開く
 * 3. デフォルトの Code.gs の中身を全部消して、このファイルの内容を貼り付ける
 * 4. 上部の「デプロイ」>「新しいデプロイ」を選択
 * 5. 種類の選択で歯車アイコンから「ウェブアプリ」を選ぶ
 * 6. 「次のユーザーとして実行」= 自分
 *    「アクセスできるユーザー」= 全員
 *    にしてデプロイする
 * 7. 発行されたウェブアプリのURL(.../exec で終わるもの)を、
 *    genius-pc-screen と genius-pc-remote 両方の index.html にある
 *    BACKEND_URL の値として貼り付ける
 *
 * コードを更新した場合は、再度「デプロイ」>「デプロイを管理」>
 * 編集(鉛筆アイコン)>「新しいバージョン」で再デプロイしないと反映されない。
 */

const SHEET_NAME = "state";

function getStateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange("A1").setValue("command");
    sheet.getRange("B1").setValue("ts");
    sheet.getRange("A2").setValue("");
    sheet.getRange("B2").setValue(0);
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action;
  const sheet = getStateSheet();

  if (action === "flash") {
    sheet.getRange("A2").setValue("flash");
    sheet.getRange("B2").setValue(Date.now());
    return jsonResponse({ ok: true });
  }

  if (action === "poll") {
    const command = sheet.getRange("A2").getValue();
    const ts = sheet.getRange("B2").getValue();
    return jsonResponse({ command: String(command), ts: Number(ts) || 0 });
  }

  if (action === "ping") {
    return jsonResponse({ ok: true, now: Date.now() });
  }

  return jsonResponse({ error: "unknown action" });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
