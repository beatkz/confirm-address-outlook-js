/* global Office, console */

async function onMessageSendHandler(event) {
  console.log("bgevent_olc.js: onMessageSendHandler 開始 ");
  event.completed({
    allowEvent: true,
  });
}

async function onMessageComposeHandler(event) {
  console.log("bgevent_olc.js: onMessageComposeHandler 開始 ");
  event.completed();
}

async function onMessageFromChangedHandler(event) {
  console.log("bgevent_olc.js: onMessageFromChangedHandler 開始");
  event.completed();
}

async function onMessageAttachmentsChangedHandler(event) {
  console.log("bgevent_olc.js: onMessageAttachmentsChangedHandler 開始");
  event.completed();
}

async function onMessageRecipientsChangedHandler(event) {
  console.log("bgevent_olc.js: onMessageRecipientsChangedHandler 開始");
  event.completed();
}

// IMPORTANT: マニフェストの FunctionName と一致させる
Office.actions.associate(
  "onMessageSendHandler", 
  onMessageSendHandler);
Office.actions.associate(
  "onMessageComposeHandler", 
  onMessageComposeHandler);
Office.actions.associate(
  "onMessageFromChangedHandler", 
  onMessageFromChangedHandler);
Office.actions.associate(
  "onMessageAttachmentsChangedHandler", 
  onMessageAttachmentsChangedHandler);
Office.actions.associate(
  "onMessageRecipientsChangedHandler",  
  onMessageRecipientsChangedHandler);
