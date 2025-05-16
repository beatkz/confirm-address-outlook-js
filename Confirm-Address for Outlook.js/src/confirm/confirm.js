/* global Office, console, document */

console.log("confirm.js: スクリプトロード開始");

// Office.js の初期化
Office.onReady((info) => {
  // ここでClassic Outlookを判定
  if (info.host === Office.HostType.Outlook && info.platform === Office.PlatformType.PC) {
    console.warn("confirm.js: Outlook Classic (Win32) ではサポートされていません。処理を中断します。");
    document.body.innerHTML = "<div id='platformError'>このアドインはOutlook Classicではサポートされていません。</div>";
    return;
  }

  // ここから既存の初期化処理
  console.log("confirm.js: Office.js 初期化完了:", JSON.stringify(info));
  console.log("confirm.js: ホスト:", info.host, "プラットフォーム:", info.platform);
  console.log("confirm.js: Office.context:", Office.context);
  console.log("confirm.js: Office.context.ui:", Office.context.ui);

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    onMessageFromParent,
    onRegisterMessageComplete
  );
  Office.context.ui.messageParent(JSON.stringify({ type: "dialogReady" }));
  console.log("confirm.js: 準備完了メッセージを送信");
}).catch((error) => {
  console.error("confirm.js: Office.js 初期化エラー:", error);
});

// メッセージを処理
function onMessageFromParent(recv) {
  try {
    const emailDetails = JSON.parse(recv.message);
    console.log("confirm.js: メール詳細を受信:", emailDetails);
    document.getElementById("toReci").textContent = emailDetails.toReci;
    document.getElementById("ccReci").textContent = emailDetails.ccReci;
    document.getElementById("bccReci").textContent = emailDetails.bccReci;
    document.getElementById("body").textContent = emailDetails.body;
    document.getElementById("attNames").textContent = emailDetails.attNames;
  } catch (error) {
    console.error("confirm.js: メール詳細の解析エラー:", error);
  }
}

function onRegisterMessageComplete(result) {
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.error("confirm.js: addHandlerAsync エラー:", result.error.message);
  } else {
    console.log("confirm.js: DialogParentMessageReceived ハンドラ登録成功");
  }
}

function cancelSend() {
  console.log("confirm.js: cancelSend 実行");
  const msgTo = { type: "cancel" };
  Office.context.ui.messageParent(JSON.stringify(msgTo));
}

function confirmSend() {
  // 送信処理をここに追加
  console.log("confirm.js: confirmSend 実行");
  const msgTo = { type: "confirm" };
  Office.context.ui.messageParent(JSON.stringify(msgTo));
}
