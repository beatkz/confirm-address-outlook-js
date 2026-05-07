/* global Office, console */

// Outlook Classic版限定 簡易アドレス確認
// 送信前に宛先をダイアログメッセージで確認（SoftBlock使用）
async function onMessageSendHandler(event) {
  console.log("bgevent_olc.js: Outlook Classic用簡易アドレス確認 開始");

  try {
    const item = Office.context.mailbox.item;
    const recipients = await collectRecipients(item);
    
    if (recipients.length === 0) {
      event.completed({ allowEvent: true });
      return;
    }

    const addrList = recipients.join("\n");
    const confirmMsg = `【アドレス確認】\n\n以下のアドレスに送信します。\n${addrList}\n\nこの内容で送信してよろしいですか？\n\n※Confirm-Address for Outlook Classicがインストールされている場合、続けて送信確認ダイアログが表示されます。`;

    event.completed({
      allowEvent: false,
      errorMessage: confirmMsg
    });
  } catch (error) {
    console.error("bgevent_olc.js: エラー発生:", error);
    event.completed({ allowEvent: true });
  }
}

async function collectRecipients(item) {
  const recipients = [];
  const fields = [
    { name: "to", label: "To" },
    { name: "cc", label: "Cc" },
    { name: "bcc", label: "Bcc" }
  ];

  for (const field of fields) {
    if (item[field.name] && typeof item[field.name].getAsync === "function") {
      try {
        await new Promise((resolve) => {
          item[field.name].getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value) {
              const adds = asyncResult.value
                .filter(r => r && r.emailAddress)
                .map(r => `${field.label}: ${r.emailAddress}`);
              recipients.push(...adds);
            }
            resolve();
          });
        });
      } catch (e) {
        console.warn(`bgevent_olc.js: ${field.name} 取得エラー`, e);
      }
    }
  }
  console.log("bgevent_olc.js: 収集したアドレス:", recipients);
  return recipients;
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);