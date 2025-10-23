/* global Office, console */

console.log("bgevent.js: スクリプトロード開始");

// HTMLタグを除去してプレーンテキストに変換する関数（改行を保持）
function stripHtml(html) {
  try {
    // HTMLエスケープされた文字列をデコード
    const textarea = document.createElement("textarea");
    textarea.innerHTML = html;
    const decodedHtml = textarea.value;
    console.log("bgevent.js: stripHtml 入力:", decodedHtml);

    // 一時的なdiv要素を作成
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = decodedHtml;

    // テキストノードと改行タグを再帰的に処理
    function extractTextWithBreaks(node, result = []) {
      for (const child of node.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) {
          // テキストノードを追加
          let text = child.textContent.trim();
          if (text) {
            result.push(text);
          }
        } else if (child.nodeType === Node.ELEMENT_NODE) {
          // 改行タグ（<br>, <p>, <div>）を検出して改行を追加
          const tagName = child.tagName.toLowerCase();
          if (tagName === "br" || tagName === "p" || tagName === "div") {
            result.push("\n");
          }
          extractTextWithBreaks(child, result);
        }
      }
      return result;
    }

    // テキストと改行を抽出
    const textArray = extractTextWithBreaks(tempDiv);
    // テキストを結合、連続する改行を保持し、スペースを整理
    let text = textArray.join("\n").replace(/\s+/g, " ").trim();
    console.log("bgevent.js: stripHtml 出力:", text);
    // 空の場合のフォールバック
    return text || "本文なし";
  } catch (error) {
    console.error("bgevent.js: HTMLタグ除去エラー:", error);
    return "本文なし";
  }
}

Office.onReady((info) => {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
}).catch((error) => {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

Office.context.mailbox.item.addHandlerAsync(
  Office.EventType.ItemChanged,
  onItemChanged,
  (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("bgevent.js: ItemChangedハンドラ登録エラー:", result.error.message);
    } else {
      console.log("bgevent.js: ItemChangedハンドラ登録成功");
    }
  }
);

let countDownTimer = null;

function onItemChanged() {
  console.log("bgevent.js: onItemChanged発火");
  clearTimeout(countDownTimer);
  Office.context.ui.closeContainer();
}

function checkAddress(args, retryCount = 0) {
  console.log("bgevent.js: checkAddress実行");
  if (!Office.context.mailbox.item.to) {
    console.warn("bgevent.js: 宛先が未定義、処理を中断");
    Office.context.ui.closeContainer();
    return;
  }

  const prefs = Office.context.roamingSettings;
  const insiderDomains = prefs.get("insiderDomains")?.split(",").map((d) => d.trim().toLowerCase()) || [];
  const maxBodyLines = parseInt(prefs.get("confirmMailBodyLines") || 5, 10);

  const emailDetails = {
    insiderReci: [],
    outsiderReci: [],
    attNames: [],
    body: "",
  };

  const msgCompFields = Office.context.mailbox.item;

  function resolveEmailAddressesAsync(addType, resolve) {
    msgCompFields[addType].getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(`bgevent.js: ${addType} 取得エラー:`, result.error.message);
        resolve([]);
        return;
      }
      const recipients = result.value.map((recipient) => ({
        type: addType === "to" ? "To: " : addType === "cc" ? "Cc: " : "Bcc: ",
        address: recipient.emailAddress.toLowerCase(),
      }));
      resolve(recipients);
    });
  }

  function resolveAttachmentsAsync(resolve) {
    if (!msgCompFields.attachments) {
      resolve([]);
      return;
    }
    const attachments = msgCompFields.attachments.map((att) => ({
      name: att.name,
    }));
    resolve(attachments);
  }

  function resolveBodyAsync(resolve) {
    msgCompFields.body.getAsync(Office.CoercionType.Html, (htmlResult) => {
      if (htmlResult.status === Office.AsyncResultStatus.Failed) {
        console.error("bgevent.js: HTML本文取得エラー:", htmlResult.error.message);
        // テキスト形式でリトライ
        msgCompFields.body.getAsync(Office.CoercionType.Text, (textResult) => {
          if (textResult.status === Office.AsyncResultStatus.Failed) {
            console.error("bgevent.js: テキスト本文取得エラー:", textResult.error.message);
            resolve("本文なし");
          } else {
            console.log("bgevent.js: テキスト本文取得成功");
            // テキストメールはそのまま使用
            let body = textResult.value || "本文なし";
            if (prefs.get("confirmMailBody")) {
              body = body.split("\n").slice(0, maxBodyLines).join("\n").trim();
            } else {
              body = "";
            }
            resolve(body);
          }
        });
      } else {
        console.log("bgevent.js: HTML本文取得成功");
        // HTMLメールをプレーンテキストに変換
        let body = stripHtml(htmlResult.value) || "本文なし";
        if (prefs.get("confirmMailBody")) {
          body = body.split("\n").slice(0, maxBodyLines).join("\n").trim();
        } else {
          body = "";
        }
        resolve(body);
      }
    });
  }

  Promise.all([
    new Promise((resolve) => resolveEmailAddressesAsync("to", resolve)),
    new Promise((resolve) => resolveEmailAddressesAsync("cc", resolve)),
    new Promise((resolve) => resolveEmailAddressesAsync("bcc", resolve)),
    new Promise((resolve) => resolveAttachmentsAsync(resolve)),
    new Promise((resolve) => resolveBodyAsync(resolve)),
  ])
    .then(([toList, ccList, bccList, attList, body]) => {
      console.log("bgevent.js: データ取得完了");
      const allReci = [...toList, ...ccList, ...bccList];
      allReci.forEach((recipient) => {
        const domain = recipient.address.split("@")[1]?.toLowerCase();
        if (insiderDomains.includes(domain)) {
          emailDetails.insiderReci.push(recipient);
        } else {
          emailDetails.outsiderReci.push(recipient);
        }
      });
      emailDetails.attNames = attList;
      emailDetails.body = body;

      Office.context.ui.displayDialogAsync(
        "https://localhost:3000/capopup.html",
        { height: 50, width: 30, displayInIframe: true },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("bgevent.js: ダイアログ表示エラー:", result.error.message);
            Office.context.ui.closeContainer();
            return;
          }
          const dialog = result.value;
          console.log("bgevent.js: ダイアログ表示成功");
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (recv) => {
            const message = JSON.parse(recv.message);
            console.log("bgevent.js: ダイアログメッセージ受信:", message);
            if (message.type === "dialogReady") {
              dialog.messageChild(JSON.stringify(emailDetails));
            } else if (message.type === "confirm") {
              dialog.close();
              args.completed({ allowEvent: true });
            } else if (message.type === "countDown") {
              let seconds = message.seconds;
              countDownTimer = setInterval(() => {
                seconds--;
                if (seconds <= 0) {
                  clearInterval(countDownTimer);
                  dialog.close();
                  args.completed({ allowEvent: true });
                } else {
                  dialog.messageChild(JSON.stringify({ type: "countdownUpdate", seconds }));
                }
              }, 1000);
            } else if (message.type === "cancel") {
              clearInterval(countDownTimer);
              dialog.close();
              args.completed({ allowEvent: false });
            }
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (recv) => {
            console.error("bgevent.js: ダイアログエラー:", JSON.stringify(recv));
            clearInterval(countDownTimer);
            dialog.close();
            args.completed({ allowEvent: false });
          });
        }
      );
    })
    .catch((error) => {
      console.error("bgevent.js: データ取得エラー:", error);
      if (retryCount < 3) {
        console.log("bgevent.js: リトライ", retryCount + 1);
        setTimeout(() => checkAddress(args, retryCount + 1), 1000);
      } else {
        console.error("bgevent.js: リトライ上限到達、処理中断");
        args.completed({ allowEvent: false });
      }
    });
}

function onMessageSend(args) {
  console.log("bgevent.js: onMessageSend発火");
  clearTimeout(countDownTimer);
  checkAddress(args);
}