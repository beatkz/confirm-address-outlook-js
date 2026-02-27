/* global Office, console */

// メインのイベントハンドラー(進行バグ回避用ダミー実装)
// 実際の処理はConfirm-Address for Outlook Classicに任せる
function onMessageSendHandler(event) {
  event.completed({ allowEvent: true });
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);