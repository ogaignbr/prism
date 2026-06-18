/**
 * PRISM 注文同期スクリプト（Google Apps Script）
 * Snipcart の Webhook を受け取り、注文をスプレッドシートに1行ずつ追記する。
 *
 * ─────────────────────────────────────────────
 * 【セットアップ手順】
 * 1. Google ドライブで新しいスプレッドシートを作成し、URL の
 *    /d/ と /edit の間の文字列（スプレッドシートID）を下の SHEET_ID に貼る。
 * 2. このスクリプトを「拡張機能 → Apps Script」に貼り付ける。
 * 3. SECRET_TOKEN を推測されにくい任意の文字列に変更する（例：ランダムな英数字）。
 * 4. 「デプロイ → 新しいデプロイ → 種類:ウェブアプリ」
 *      - 実行ユーザー：自分
 *      - アクセスできるユーザー：全員
 *    を選び、発行された Web アプリ URL を控える。
 * 5. その URL の末尾に ?token=（SECRET_TOKENと同じ値） を付けたものを、
 *    Snipcart ダッシュボード → Settings → Webhooks の「Your URL」に登録する。
 *      例： https://script.google.com/macros/s/AKfy.../exec?token=あなたの秘密トークン
 * 6. 対象イベントは order.completed（必要に応じて order.status.changed など）。
 *
 * ※ Apps Script の doPost はリクエストヘッダーを取得できないため、
 *   Snipcart 標準の署名検証は使えない。代わりに上記 ?token= による簡易認証を行う。
 * ─────────────────────────────────────────────
 */

var SHEET_ID = "ここにスプレッドシートIDを貼る";
var SHEET_NAME = "orders"; // タブ名（なければ自動作成）
var SECRET_TOKEN = "ここに秘密トークンを設定する";

// スプレッドシートの見出し行
var HEADERS = [
  "受信日時",
  "イベント",
  "注文番号",
  "氏名",
  "メール",
  "商品明細",
  "合計",
  "通貨",
  "支払方法",
  "支払状態",
  "配送先",
  "電話",
];

function doPost(e) {
  try {
    // 簡易認証：?token= が一致しない場合は拒否
    if (!e || !e.parameter || e.parameter.token !== SECRET_TOKEN) {
      return jsonResponse({ ok: false, error: "unauthorized" });
    }
    if (!e.postData || !e.postData.contents) {
      return jsonResponse({ ok: true, note: "no payload" });
    }

    var payload = JSON.parse(e.postData.contents);
    var eventName = payload.eventName || "";
    var order = payload.content || {};

    // 注文系イベントのみ記録
    if (eventName.indexOf("order") === 0) {
      appendOrderRow(eventName, order);
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function appendOrderRow(eventName, order) {
  var sheet = getSheet();

  var items = (order.items || [])
    .map(function (it) {
      return it.name + " x" + it.quantity;
    })
    .join(" / ");

  var ship = order.shippingAddress || order.billingAddress || {};
  var shipText = [ship.fullAddress, ship.city, ship.postalCode, ship.country]
    .filter(function (v) {
      return v;
    })
    .join(" ");

  var row = [
    new Date(),
    eventName,
    order.invoiceNumber || order.token || "",
    order.billingAddressName || order.cardHolderName || ship.name || "",
    order.email || "",
    items,
    order.finalGrandTotal || order.total || "",
    order.currency || "",
    order.paymentMethod || "",
    order.paymentStatus || order.status || "",
    shipText,
    ship.phone || order.billingAddressPhone || "",
  ];

  sheet.appendRow(row);
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // 見出し行がなければ追加
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// 動作確認用：Apps Script エディタから手動実行すると、テスト行が1行追加される
function testAppend() {
  appendOrderRow("order.completed", {
    invoiceNumber: "TEST-0001",
    billingAddressName: "山田 花子",
    email: "test@example.com",
    items: [{ name: "プリズム シャンプー ブルー（450mL）", quantity: 1 }],
    total: 4860,
    currency: "jpy",
    paymentMethod: "CreditCard",
    paymentStatus: "Paid",
    shippingAddress: {
      fullAddress: "テスト区テスト1-2-3",
      city: "東京都",
      postalCode: "100-0001",
      country: "JP",
      phone: "090-0000-0000",
    },
  });
}
