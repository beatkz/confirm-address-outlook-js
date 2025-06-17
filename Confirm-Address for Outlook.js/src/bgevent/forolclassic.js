/* global Office, console, document, setInterval, clearInterval, process */

Office.onReady((info) => {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("uniqueMessageSendHandler", uniqueMessageSendHandler);
}).catch((error) => {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

// メインのイベントハンドラ
function uniqueMessageSendHandler(event) {
  console.log("bgevent.js: uniqueMessageSendHandler 開始");
  passingThruEvent(event);
  console.log("bgevent.js: Outlook Classicではダイアログを表示せず、Confirm-Address for Outlook Classicに渡します。");
}

function passingThruEvent(event) {
  console.log("bgevent.js: passingThruEvent 開始");
  event.completed({ allowEvent: true });
}