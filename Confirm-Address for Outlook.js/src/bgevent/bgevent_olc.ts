/* global Office, console */

// 事前処理個別ハンドラー (マニフェストFunctionName準拠)
async function onMessageSendHandler(event: Office.AddinCommands.Event) {
  onPreProcessHandler(event, "send");
}

// 事前処理共通ハンドラー (Compose/From/Attachments/Recipients用)
async function onPreProcessHandler(event: Office.AddinCommands.Event, eventType: string) {
  console.log(`bgevent.js: onPreProcessHandler (${eventType}) 開始`);
  event.completed({ allowEvent: true });
}

// IMPORTANT: マニフェストの FunctionName と一致させる
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
