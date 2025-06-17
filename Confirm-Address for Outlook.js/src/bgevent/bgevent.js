/* global Office, console, setInterval, clearInterval, process */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook && info.platform === Office.PlatformType.PC) {
    isOLClassic = true;
  }

  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("uniqueMessageSendHandler", uniqueMessageSendHandler);
}).catch((error) => {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

// ダイアログを表示
let isOLClassic = false; // [Outlook Classic向け]ダイアログを抑制するフラグ
let caDialog; // confirmダイアログのグローバル変数
let countdownInterval = null; // カウントダウン用のグローバル変数

// メインのイベントハンドラ
function uniqueMessageSendHandler(event) {
  console.log("bgevent.js: uniqueMessageSendHandler 開始");
  if (isOLClassic) {
    passingThruEvent(event);
    console.log("bgevent.js: Outlook Classicではダイアログを表示せず、Confirm-Address for Outlook Classicに渡します。");
    return;
  }
  showConfirmDialog(event);
}

function passingThruEvent(event) {
  console.log("bgevent.js: passingThruEvent 開始");
  event.completed({ allowEvent: true });
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
  const bodyResult = await new Promise((resolve) => msgCompFields.body.getAsync("text", resolve));
  const lines = Office.context.roamingSettings.get("confirmMailBodyLines") || 5;
  const body = bodyResult.value.split("\n").slice(0, lines).join("\n") || "本文なし";

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
