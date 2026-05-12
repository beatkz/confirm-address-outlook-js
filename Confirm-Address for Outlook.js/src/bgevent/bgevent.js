/* global Office, console, setInterval, clearInterval, process, document, Node */

let caDialog; // confirmダイアログのグローバル変数
let countdownInterval = null; // カウントダウン用のグローバル変数

Office.onReady((info) => {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  // 事前処理イベントはすべて同一の共通関数を使用（重複排除）
  Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
  Office.actions.associate("onMessageFromChangedHandler", onMessageFromChangedHandler);
  Office.actions.associate(
    "onMessageAttachmentsChangedHandler",
    onMessageAttachmentsChangedHandler
  );
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
}).catch((error) => {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

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
    console.log("bgevent.js: stripHtml 出力:", text);
    return text || "本文なし";
  } catch (error) {
    console.error("bgevent.js: HTMLタグ除去エラー:", error);
    return "本文なし";
  }
}

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

// 事前処理用のイベントハンドラー

function onMessageComposeHandler(event) {
  console.log("bgevent.js: onMessageComposeHandler 開始 - キャッシュ更新");
  updateCacheData();
  event.completed();
}

function onMessageFromChangedHandler(event) {
  console.log("bgevent.js: onMessageFromChangedHandler 開始 - キャッシュ更新");
  updateCacheData();
  event.completed();
}

function onMessageAttachmentsChangedHandler(event) {
  console.log("bgevent.js: onMessageAttachmentsChangedHandler 開始 - キャッシュ更新");
  updateCacheData();
  event.completed();
}

function onMessageRecipientsChangedHandler(event) {
  console.log("bgevent.js: onMessageRecipientsChangedHandler 開始 - キャッシュ更新");
  updateCacheData();
  event.completed();
}

// キャッシュデータをカスタムプロパティからロード
async function loadCacheDetails() {
  console.log("bgevent.js: loadCacheDetails 開始");
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const customProps = result.value;
        const detailsStr = customProps.get("caCacheDetails");
        if (detailsStr) {
          try {
            const details = JSON.parse(detailsStr);
            console.log("bgevent.js: Cache loaded");
            resolve(details);
            return;
          } catch (e) {
            console.error("bgevent.js: Cache details parse error:", e);
          }
        }
      } else {
        console.error("bgevent.js: loadCustomPropertiesAsync failed:", result.error);
      }
      resolve(getDefaultCache());
    });
  });
}

function getDefaultCache() {
  return {
    senderAddress: "",
    caReciList: {
      insider: [],
      outsider: [],
    },
    attNames: [],
  };
}

// キャッシュを保存
function saveCacheDetails(details) {
  console.log("bgevent.js: saveCacheDetails 開始");
  const item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const customProps = result.value;
      customProps.set("caCacheDetails", JSON.stringify(details));
      customProps.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("bgevent.js: Cache details saved");
        } else {
          console.error("bgevent.js: saveCustomPropertiesAsync failed:", saveResult.error);
        }
      });
    } else {
      console.error("bgevent.js: load for save failed:", result.error);
    }
  });
}

// キャッシュデータ計算（本文以外）
async function fetchCacheData() {
  console.log("bgevent.js: fetchCacheData 開始");
  const msgCompFields = Office.context.mailbox.item;

  const senderAddress = await getSenderAddress(msgCompFields);
  console.log("bgevent.js: 送信者アドレス:", senderAddress);

  var reciList = {
    to: [],
    cc: [],
    bcc: [],
  };

  await collectAddress(msgCompFields, reciList);
  console.log("bgevent.js: メールアドレス収集完了");
  console.dir("bgevent.js: ", reciList);

  var domainList = getDomainList();
  console.log("bgevent.js: 組織のドメインリスト:", domainList);

  var caReciList = {
    insider: [],
    outsider: [],
  };

  for (const reciType of ["to", "cc", "bcc"]) {
    judgeAddress(reciList[reciType], domainList, caReciList);
  }

  console.log("bgevent.js: 組織内アドレス:", caReciList.insider.map((r) => r.address).join(", "));
  console.log("bgevent.js: 組織外アドレス:", caReciList.outsider.map((r) => r.address).join(", "));

  var attNames = [];
  await getAttachments(msgCompFields, attNames);
  console.log("bgevent.js: 添付ファイル名:", attNames.map((att) => att.name).join(", "));

  return {
    senderAddress: senderAddress,
    caReciList: caReciList,
    attNames: attNames,
  };
}

// 更新トリガー
function updateCacheData() {
  fetchCacheData()
    .then((details) => {
      saveCacheDetails(details);
    })
    .catch((error) => {
      console.error("bgevent.js: fetchCacheData error:", error);
    });
}

async function getMailBody(msgCompFields) {
  let body;
  const lines = Office.context.roamingSettings.get("confirmMailBodyLines") || 5;
  const htmlResult = await new Promise((resolve) => msgCompFields.body.getAsync("html", resolve));
  console.log("bgevent.js: HTML本文取得完了");
  const rawBody = htmlResult.value;
  const hasHtmlTags = /<\/?[a-z][^>]*>/i.test(rawBody);
  console.log("bgevent.js: HTMLタグ検知:", hasHtmlTags);
  if (hasHtmlTags) {
    body = stripHtml(rawBody);
  } else {
    body = rawBody;
  }
  if (!body) {
    body = "本文なし";
  }
  if (Office.context.roamingSettings.get("confirmMailBody")) {
    body = body.split("\n").slice(0, lines).join("\n").trim();
  } else {
    body = "";
  }
  return body;
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

async function collectAddress(msgCompFields, reciList) {
  console.log("bgevent.js: collectAddress 開始");
  await collectFieldRecipients(msgCompFields, "to", reciList.to, "To: ");
  await collectFieldRecipients(msgCompFields, "cc", reciList.cc, "Cc: ");
  await collectFieldRecipients(msgCompFields, "bcc", reciList.bcc, "Bcc: ");
}

async function getAttachments(msgCompFields, attList) {
  const attResult = await new Promise((resolve) => msgCompFields.getAttachmentsAsync(resolve));
  const tempAtt = attResult.value.map((att) => att.name) || [];
  for (const att of tempAtt) {
    if (att) {
      attList.push({ name: att });
    }
  }
}

// メール詳細をダイアログに送信（事前計算データ＋本文）
async function sendEmailDetails() {
  console.log("bgevent.js: sendEmailDetails 開始");
  const Cache = await loadCacheDetails();
  const msgCompFields = Office.context.mailbox.item;
  let body = await getMailBody(msgCompFields);
  const emailDetails = {
    ...Cache,
    body: body,
  };
  console.log("bgevent.js: 詳細:", emailDetails);
  caDialog.messageChild(JSON.stringify(emailDetails));
}
