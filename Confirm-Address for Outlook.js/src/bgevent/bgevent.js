/* global Office, console, setInterval, clearInterval, process */

// ダイアログを表示
let caDialog; // confirmダイアログのグローバル変数
let countdownInterval = null; // カウントダウン用のグローバル変数

Office.onReady((info) => {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
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
  switch (msgDlg.type) {
    case "dialogReady":
      console.log("bgevent.js: ダイアログ準備完了メッセージを受信、メール詳細を送信");
      sendEmailDetails();
      break;
    case "confirm":
      console.log("bgevent.js: 確認メッセージを受信、送信を許可");
      dialog.close();
      sendEvent.completed({ allowEvent: true });
      if (countdownInterval) {
        clearInterval(countdownInterval);
        countdownInterval = null;
      }
      break;
    case "countDown":
      console.log("bgevent.js: カウントダウン開始:", msgDlg.seconds);
      startCountdown(msgDlg.seconds, sendEvent, dialog);
      break;
    case "cancel":
      console.log("bgevent.js: キャンセルメッセージを受信、送信をキャンセル");
      dialog.close();
      sendEvent.completed({
        allowEvent: false,
        errorMessage: "送信がキャンセルされました。",
      });
      break;
    default:
      console.warn("bgevent.js: 無効なメッセージを無視:", recv, "タイプ:", typeof recv);
  }
}

function startCountdown(seconds, sendEvent, dialog) {
  let remaining = seconds;
  countdownInterval = setInterval(() => {
    if (remaining <= 0) {
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

  var toList = [];
  var ccList = [];
  var bccList = [];
  await collectAddress(msgCompFields, toList, ccList, bccList);
  console.log("bgevent.js: メールアドレス収集完了");
  console.log("bgevent.js: To:", toList, "Cc:", ccList, "Bcc:", bccList);

  var domainList = getDomainList(); // 組織のドメインリスト
  console.log("bgevent.js: 組織のドメインリスト:", domainList);

  var insiderReci = [];
  var outsiderReci = [];
  judgeAddress(toList, domainList, insiderReci, outsiderReci);
  judgeAddress(ccList, domainList, insiderReci, outsiderReci);
  judgeAddress(bccList, domainList, insiderReci, outsiderReci);
  console.log("bgevent.js: 組織内アドレス:", insiderReci.map((r) => r.address).join(", "));
  console.log("bgevent.js: 組織外アドレス:", outsiderReci.map((r) => r.address).join(", "));

  // 本文冒頭
  const lines = Office.context.roamingSettings.get("confirmMailBodyLines") || 5;
  let body;
  const htmlResult = await new Promise((resolve) => msgCompFields.body.getAsync("html", resolve));
  if (htmlResult.status === Office.AsyncResultStatus.Failed) {
    console.error("bgevent.js: HTML本文取得エラー:", htmlResult.error.message);
    const textResult = await new Promise((resolve) => msgCompFields.body.getAsync("text", resolve));
    if (textResult.status === Office.AsyncResultStatus.Failed) {
      console.error("bgevent.js: テキスト本文取得エラー:", textResult.error.message);
      body = "本文なし";
    } else {
      console.log("bgevent.js: テキスト本文取得成功");
      body = textResult.value || "本文なし";
      body = slicedMailBodyfromRawMail(body, lines);
    }
  } else {
    console.log("bgevent.js: HTML本文取得成功");
    body = stripHtml(htmlResult.value) || "本文なし";
    body = slicedMailBodyfromRawMail(body, lines);
  }

  function slicedMailBodyfromRawMail(rawmail, lines){
    let sMailBody;
    if (Office.context.roamingSettings.get("confirmMailBody")) {
      sMailbody = rawmail.split("\n").slice(0, lines).join("\n").trim();
    } else {
      sMailbody = "";
    }
    return sMailBody;
  }

  var attNames = [];
  await getAttachments(msgCompFields, attNames);
  console.log("bgevent.js: 添付ファイル名:", attNames.map((att) => att.name).join(", "));

  return {
    insiderReci: insiderReci,
    outsiderReci: outsiderReci,
    body: body,
    attNames: attNames,
  };
}

function judgeAddress(addressArray, domainList, insiderAddress, outsiderAddress) {
  console.log("bgevent.js: judgeAddress 開始");
  console.log("[JUDGE] " + addressArray.map((a) => a.address).join(", ") + "\n");

  // domainListが空の場合、全て外部とみなす
  if (domainList.length === 0) {
    for (const address of addressArray) {
      outsiderAddress.push(address);
    }
    return;
  }

  // 登録されたドメインリストとアドレスを比較
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
      insiderAddress.push(target);
    } else {
      outsiderAddress.push(target);
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

async function collectAddress(msgCompFields, toList, ccList, bccList) {
  console.log("bgevent.js: collectAddress 開始");

  // To
  const toMap = await new Promise((resolve) => msgCompFields.to.getAsync(resolve));
  const tempTo = toMap.value.map((r) => r.emailAddress) || [];
  for (const reci of tempTo) {
    if (reci) {
      toList.push({ type: "To: ", address: reci });
    }
  }

  // Cc
  const ccMap = await new Promise((resolve) => msgCompFields.cc.getAsync(resolve));
  const tempCc = ccMap.value.map((r) => r.emailAddress) || [];
  for (const reci of tempCc) {
    if (reci) {
      ccList.push({ type: "Cc: ", address: reci });
    }
  }

  // Bcc
  const bccMap = await new Promise((resolve) => msgCompFields.bcc.getAsync(resolve));
  const tempBcc = bccMap.value.map((r) => r.emailAddress) || [];
  for (const reci of tempBcc) {
    if (reci) {
      bccList.push({ type: "Bcc: ", address: reci });
    }
  }
}

async function getAttachments(msgCompFields, attList) {
  // 添付ファイル名
  const attResult = await new Promise((resolve) => msgCompFields.getAttachmentsAsync(resolve));
  const tempAtt = attResult.value.map((att) => att.name) || [];
  for (const att of tempAtt) {
    if (att) {
      attList.push({ name: att });
    }
  }
}

async function sendEmailDetails() {
  console.log("bgevent.js: sendEmailDetails 開始");

  // メールの詳細を収集する処理をここに追加
  const emailDetails = await checkAddress();

  console.log("bgevent.js: 詳細:", emailDetails);
  caDialog.messageChild(JSON.stringify(emailDetails));
}
