/* global Office, console */

// Outlook Classic版ではVSTO版で処理するため、パススルー処理します。
async function onMessageSendHandler(event) {
  console.log("bgevent_olc.js: Outlook Classic用パススルー 開始");

  try {
    event.completed({
      allowEvent: true
    });
  } catch (error) {
    console.error("bgevent_olc.js: エラー発生:", error);
    event.completed({ allowEvent: true });
  }
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
