/******/ (function() { // webpackBootstrap
/*!**********************************!*\
  !*** ./src/bgevent/bgevent_c.js ***!
  \**********************************/
/* global Office, console, setInterval, clearInterval, process */

// Outlook Classic用ダミーコード

Office.onReady(function (info) {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}).catch(function (error) {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

// Outlookのバージョンを確認して、Classicかどうかを判定する
function isOutlookClassicVersion() {
  var isOutlookClassic = Office.context.mailbox.diagnostics.hostName === "Outlook";
  var isWindowsPlatform = Office.context.platform === Office.PlatformType.PC;
  var outlookVersion = Office.context.mailbox.diagnostics.hostVersion;
  return isOutlookClassic && isWindowsPlatform && parseFloat(outlookVersion) < 16;
}

// メインのイベントハンドラ
function onMessageSendHandler(event) {
  console.log("bgevent.js: onMessageSendHandler 開始");
  if (isOutlookClassicVersion()) {
    console.log("bgevent.js: Outlook Classicを検出、ダイアログをバイパス");
    event.completed({
      allowEvent: true
    });
  }
}
/******/ })()
;
//# sourceMappingURL=bgevent_c.js.map