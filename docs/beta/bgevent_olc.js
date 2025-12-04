/******/ (function() { // webpackBootstrap
/*!************************************!*\
  !*** ./src/bgevent/bgevent_olc.js ***!
  \************************************/
/* global Office, console, setInterval, clearInterval, process */

Office.onReady(function (info) {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}).catch(function (error) {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

// メインのイベントハンドラ
function onMessageSendHandler(event) {
  console.log("bgevent.js: onMessageSendHandler 開始");
  passThruDialog(event);
}
function passThruDialog(sendEvent) {
  sendEvent.completed({
    allowEvent: true
  });
}
/******/ })()
;
//# sourceMappingURL=bgevent_olc.js.map