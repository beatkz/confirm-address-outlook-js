/* global Office, console, setInterval, clearInterval, process, document, Node */

let caDialog; // confirmダイアログのグローバル変数
let countdownInterval = null; // カウントダウン用のグローバル変数

Office.onReady((info) => {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}).catch((error) => {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

// メインのイベントハンドラ
function onMessageSendHandler(event) {
  console.log("bgevent.js: onMessageSendHandler 開始");
  showConfirmDialog(event);
}

function showConfirmDialog(sendEvent) {
  const dialogUrl = `${process.env.BASE_URL}capopup.html`;
  console.log("bgevent.js: ダイアログ表示を試行", dialogUrl);

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {
      height: 50,
      width: 30,
      promptBeforeOpen: false,
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("bgevent.js: ダイアログ表示エラー:", result.error.message);
        sendEvent.completed({
          allowEvent: false,
          errorMessage: `確認画面の表示に失敗: ${result.error.message}`,
        });
        return;
      }

      console.log("bgevent.js: ダイアログ表示成功");
      caDialog = result.value;
      console.log("bgevent.js: caDialog オブジェクト:", caDialog, "タイプ:", typeof caDialog);

      // Office.js のイベントハンドラ
      caDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        console.log("bgevent.js: ダイアログからのメッセージ受信:", arg.message);
        handleMessage(arg.message, sendEvent, caDialog);
      });

      caDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        console.log("bgevent.js: ダイアログが閉じられました:", arg);
      });
    }
  );
}

// メッセージを処理
function handleMessage(recv, sendEvent, dialog) {
  const msgDlg = JSON.parse(recv);
  const msgFuncs = {
    dialogReady: function () {
      console.log("bgevent.js: ダイアログ準備完了メッセージを受信、メール詳細を送信");
      sendEmailDetails();
    },
    confirm: function () {
      console.log("bgevent.js: 確認メッセージを受信、送信を許可");
      dialog.close();
      sendEvent.completed({ allowEvent: true });
      if (countdownInterval) {
        clearInterval(countdownInterval);
        countdownInterval = null;
      }
    },
    countDown: function () {
      console.log("bgevent.js: カウントダウン開始:", msgDlg.seconds);
      startCountdown(msgDlg.seconds, sendEvent, dialog);
    },
    cancel: function () {
      console.log("bgevent.js: キャンセルメッセージを受信、送信をキャンセル");
      dialog.close();
      sendEvent.completed({
        allowEvent: false,
        errorMessage: "送信がキャンセルされました。",
      });
    },
  };
  msgFuncs[msgDlg.type]
    ? msgFuncs[msgDlg.type]()
    : console.warn("bgevent.js: 無効なメッセージを無視:", recv, "タイプ:", typeof recv);
}

function startCountdown(seconds, sendEvent, dialog) {
  let remaining = seconds - 1;
  countdownInterval = setInterval(() => {
    if (remaining < 0) {
      console.log("bgevent.js: カウントダウン終了、送信を許可");
      clearInterval(countdownInterval);
      countdownInterval = null;
      dialog.close();
      sendEvent.completed({ allowEvent: true });
      return;
    }
    console.log("bgevent.js: カウントダウン残り:", remaining);
    dialog.messageChild(JSON.stringify({ type: "countdownUpdate", seconds: remaining }));
    remaining--;
  }, 1000);
}

async function checkAddress() {
  console.log("bgevent.js: checkAddress 開始");
  const msgCompFields = Office.context.mailbox.item;

  const senderAddress = await getSenderAddress(msgCompFields);
  console.log("bgevent.js: 送信者アドレス:", senderAddress);

  const caReciList = await caReciListfromRecipients(msgCompFields);

  // 本文冒頭
  const body = await getEmailBody(msgCompFields);
  console.log("bgevent.js: 本文冒頭:", body);

  // 添付ファイル名の収集
  const attNames = await getAttachments(msgCompFields);
  console.log("bgevent.js: 添付ファイル名:", attNames.map((att) => att.name).join(", "));

  return {
    senderAddress: senderAddress,
    caReciList: caReciList,
    body: body,
    attNames: attNames,
  };
}

async function caReciListfromRecipients(msgCompFields) {
  const caReciList = {
    insider: [],
    outsider: [],
  };

  const reciList = await collectAddress(msgCompFields);
  console.log("bgevent.js: メールアドレス収集完了");
  console.dir("bgevent.js: ", reciList);

  const domainList = getDomainList(); // 組織のドメインリスト
  console.log("bgevent.js: 組織のドメインリスト:", domainList);

  for (const reciType of ["to", "cc", "bcc"]) {
    judgeAddress(reciList[reciType], domainList, caReciList);
  }
  console.dir("bgevent.js: 振分アドレス:", caReciList);
  return caReciList;
}

function judgeAddress(addressArray, domainList, caReciList) {
  console.log("bgevent.js: judgeAddress 開始");
  console.log("[JUDGE] " + addressArray.map((a) => a.address).join(", ") + "\n");

  if (domainList.length === 0) {
    for (const address of addressArray) {
      caReciList.outsider.push(address);
    }
    return;
  }

  for (const target of addressArray) {
    const address = target.address;
    if (address.length === 0) {
      continue;
    }
    const domain = address.substring(address.indexOf("@")).toLowerCase();

    let match = false;
    for (const insiderDomain of domainList) {
      if (domain.includes(insiderDomain.toLowerCase())) {
        match = true;
        break;
      }
    }
    if (match) {
      caReciList.insider.push(target);
    } else {
      caReciList.outsider.push(target);
    }
  }
}

function getDomainList() {
  console.log("bgevent.js: getDomainList 開始");
  const domainList = [];
  const settings = Office.context.roamingSettings;
  const insiderDomains = settings.get("insiderDomains");
  if (insiderDomains) {
    const domains = insiderDomains.split(",").map((domain) => domain.trim());
    for (const domain of domains) {
      if (domain) {
        domainList.push(domain);
      }
    }
  }
  return domainList;
}

async function getSenderAddress(msgCompFields) {
  console.log("bgevent.js: getSenderAddress 開始");
  const senderResult = await new Promise((resolve) => msgCompFields.from.getAsync(resolve));
  return senderResult.value.emailAddress;
}

// 受信者フィールド(To/Cc/Bcc)の収集を共通化（重複排除のためのヘルパー）
async function collectFieldRecipients(msgCompFields, fieldName, targetList, typePrefix) {
  const result = await new Promise((resolve) => msgCompFields[fieldName].getAsync(resolve));
  const addresses = result.value?.map((r) => r.emailAddress) || [];
  for (const address of addresses) {
    if (address) {
      targetList.push({ type: typePrefix, address: address });
    }
  }
}

async function collectAddress(msgCompFields) {
  var reciList = {
    to: [],
    cc: [],
    bcc: [],
  };

  console.log("bgevent.js: collectAddress 開始");
  for (const field of ["to", "cc", "bcc"]) {
    console.log(`bgevent.js: ${field}フィールドの収集を開始`);
    await collectFieldRecipients(
      msgCompFields,
      field,
      reciList[field],
      `${field.charAt(0).toUpperCase() + field.slice(1)}: `
    );
  }
  return reciList;
}

async function getEmailBody(msgCompFields) {
  const htmlResult = await new Promise((resolve) => msgCompFields.body.getAsync("html", resolve));
  console.log("bgevent.js: HTML本文取得完了");

  const rawBody = htmlResult.value;

  var body = textFromRawText(rawBody);

  if (!body) {
    return "本文なし";
  }

  if (Office.context.roamingSettings.get("confirmMailBody")) {
    const lines = Office.context.roamingSettings.get("confirmMailBodyLines") || 5;
    return body.split("\n").slice(0, lines).join("\n").trim();
  } else {
    return "";
  }
}

function detectHTMLTags(text) {
  return /<\/?[a-z][^>]*>/i.test(text);
}

// 本文をプレーンテキストに変換する関数（改行を保持）
function textFromRawText(rawText) {
  if (detectHTMLTags(rawText)) {
    try {
      // HTMLエスケープされた文字列をデコード
      const textarea = document.createElement("textarea");
      textarea.innerHTML = rawText;
      const decodedHtml = textarea.value;
      console.log("bgevent.js: textFromRawText 入力:", decodedHtml);

      // 一時的なdiv要素を作成
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = decodedHtml;

      // テキストノードと改行タグを再帰的に処理
      function extractTextWithBreaks(node, result = []) {
        for (const child of node.childNodes) {
          if (child.nodeType === Node.TEXT_NODE) {
            let text = child.textContent.trim();
            if (text) result.push(text);
          } else if (child.nodeType === Node.ELEMENT_NODE) {
            const tagName = child.tagName.toLowerCase();
            if (tagName === "br" || tagName === "p" || tagName === "div") {
              result.push("\n");
            }
            extractTextWithBreaks(child, result);
          }
        }
        return result;
      }

      const textArray = extractTextWithBreaks(tempDiv);
      console.log("textArray: ", textArray);

      let text = textArray
        .join("")
        .replace(/[ \t]+/g, " ")
        .trim();
      console.log("bgevent.js: textFromRawText 出力:", text);
      return text || "本文なし";
    } catch (error) {
      console.error("bgevent.js: HTMLタグ除去エラー:", error);
      return "本文なし";
    }
  } else {
    return rawText;
  }
}

async function getAttachments(msgCompFields) {
  const attResult = await new Promise((resolve) => msgCompFields.getAttachmentsAsync(resolve));
  const tempAtt = attResult.value.map((att) => att.name) || [];
  return tempAtt.filter((att) => !!att).map((att) => ({ name: att }));
}

// メール詳細をダイアログに送信
async function sendEmailDetails() {
  console.log("bgevent.js: sendEmailDetails 開始");
  caDialog.messageChild(JSON.stringify(await checkAddress()));
}
