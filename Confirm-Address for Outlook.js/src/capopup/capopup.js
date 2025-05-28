/* global Office, console, document */

console.log("capopup.js: スクリプトロード開始");

// Office.js の初期化
Office.onReady((info) => {
  // Classic Outlookを判定
  if (info.host === Office.HostType.Outlook && info.platform === Office.PlatformType.PC) {
    console.warn(
      "capopup.js: Outlook Classic (Win32) ではサポートされていません。処理を中断します。"
    );
    return;
  }

  // 初期化処理
  console.log("capopup.js: Office.js 初期化完了:", JSON.stringify(info));
  console.log("capopup.js: ホスト:", info.host, "プラットフォーム:", info.platform);

  // 一括チェックボタンのイベントリスナーを関数化
  const setupCheckboxListener = (checkboxId, targetId = null) => {
    const checkbox = document.getElementById(checkboxId);
    checkbox.addEventListener("change", (event) => {
      if (targetId) {
        batchCheck(targetId, event.target.checked);
      }
      checkAllChecked(); // 送信ボタンの状態を更新
    });
  };

  // チェックボックスリスナーの設定
  setupCheckboxListener("batchCheck_insiderReci", "insiderReci");
  setupCheckboxListener("batchCheck_outsiderReci", "outsiderReci");
  setupCheckboxListener("batchCheck_Attachments", "attNames");
  setupCheckboxListener("check_firstLinesOfBody", null);

  // 一括チェックボタンの有効化/無効化
  const prefs = Office.context.roamingSettings;
  const setupBatchCheckButton = (buttonId, prefId) => {
    if (prefs.get(prefId)){
      document.getElementById(buttonId).disabled = false;
    }
  };
  setupBatchCheckButton("batchCheck_insiderReci", "insiderDomainBatchCheck");
  setupBatchCheckButton("batchCheck_outsiderReci", "outsiderDomainBatchCheck");
  setupBatchCheckButton("batchCheck_Attachments", "attachmentBatchCheck");

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    onMessageFromParent,
    onRegisterMessageComplete
  );
  Office.context.ui.messageParent(JSON.stringify({ type: "dialogReady" }));
  console.log("capopup.js: 準備完了メッセージを送信");
}).catch((error) => {
  console.error("capopup.js: Office.js 初期化エラー:", error);
});

function onRegisterMessageComplete(result) {
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.error("capopup.js: addHandlerAsync エラー:", result.error.message);
  } else {
    console.log("capopup.js: DialogParentMessageReceived ハンドラ登録成功");
  }
}

function batchCheck(targetId, val) {
  const targetDiv = document.getElementById(targetId);
  const inputs = targetDiv.getElementsByTagName("input");
  for (let i = 0; i < inputs.length; i++) {
    inputs[i].checked = val;
  }
}

function handleHiddenContainers(emailDetails) {
  // 隠しコンテナの表示/非表示を切り替える
  if( emailDetails.insiderReci.length > 0) {
    const insiderContainer = document.getElementById("insiderContainer");
    insiderContainer.hidden = false;
  }
  if( emailDetails.outsiderReci.length > 0) {
    const outsiderContainer = document.getElementById("outsiderContainer");
    outsiderContainer.hidden = false;
  }
  if( emailDetails.attNames.length > 0) {
    const attNamesContainer = document.getElementById("attNamesContainer");
    attNamesContainer.hidden = false;
  }

  const prefs = Office.context.roamingSettings;
  if(prefs.get("confirmMailBody")) {
    document.getElementById("mailBodyContainer").hidden = false;
  }
}
// メッセージを処理
function onMessageFromParent(recv) {
  try {
    const message = JSON.parse(recv.message);
    console.log("capopup.js: メッセージ受信:", message);
    if (message.type === "countdownUpdate") {
      document.getElementById("countdown").textContent = `あと${message.seconds}秒で送信します。`;
      return;
    }

    const emailDetails = message;
    const prefs = Office.context.roamingSettings;
    const outsiderAddressCount = emailDetails.outsiderReci.length;
    if (prefs.get("noDisplayInsiderDomainOnly") 
      && outsiderAddressCount === 0) {
      checkAllChecked(outsiderAddressCount);
      confirmSend();
      return;
    }

    handleHiddenContainers(emailDetails);
    console.log("capopup.js: メール詳細を受信:", emailDetails);
    document.getElementById("insiderReci").textContent = "";
    pushToList({
      targetId: "insiderReci",
      listType: "Addresses",
      pushingList: emailDetails.insiderReci
    });

    document.getElementById("outsiderReci").textContent = "";
    pushToList({
      targetId: "outsiderReci",
      listType: "Addresses",
      pushingList: emailDetails.outsiderReci
    });

    document.getElementById("body").textContent = emailDetails.body;
    document.getElementById("attNames").textContent = "";
    pushToList({
      targetId: "attNames",
      listType: "Attachments",
      pushingList: emailDetails.attNames
    });

    checkAllChecked(); // 送信ボタンの状態を初期化
  } catch (error) {
    console.error("capopup.js: メール詳細の解析エラー:", error);
  }
}

function pushToList(args) {
  /*
  args:
  {
      targetId: "divfoo",
      listType: "Addresses" or "Attachments",
      pushingList: []
  }
  */
  const targetDiv = document.getElementById(args.targetId);
  let pushTxtNode;

  for (let i = 0; i < args.pushingList.length; i++) {
    switch(args.listType) {
      case "Addresses":
        pushTxtNode = args.pushingList[i].type + args.pushingList[i].address;
        break;
      case "Attachments":
        pushTxtNode = args.pushingList[i].name;
        break;
    }
    const chkbox = document.createElement("input");
    chkbox.setAttribute("type", "checkbox");
    chkbox.setAttribute("id", `checkbox-${args.targetId}-${i}`);
    chkbox.addEventListener("change", () => {
      checkAllChecked(); // 送信ボタンの状態を更新
    });
    const label = document.createElement("label");
    label.setAttribute("for", `checkbox-${args.targetId}-${i}`);
    label.appendChild(chkbox);
    label.appendChild(document.createTextNode(pushTxtNode));
    targetDiv.appendChild(label);
    targetDiv.appendChild(document.createElement("br"));
  }
}

function checkAllChecked(outsiderAddressCount) {
  // チェックボックスリストの状態を確認する汎用関数
  const areAllCheckboxesChecked = (containerId) => {
    const container = document.getElementById(containerId);
    const checkboxes = container.getElementsByTagName('input');
    return Array.from(checkboxes).every(checkbox => checkbox.checked);
  };
  const prefs = Office.context.roamingSettings;
  
  // 各セクションのチェック状態を確認
  const isInsiderDomainsChecked = areAllCheckboxesChecked('insiderReci');
  const isOutsiderDomainsChecked = areAllCheckboxesChecked('outsiderReci');
  let isMailHeadChecked = true; // confirmMailBodyが無効な場合はtrueに設定
  if(prefs.get("confirmMailBody")) {
    isMailHeadChecked = document.getElementById('check_firstLinesOfBody').checked;
  }
  const isAttachmentsChecked = areAllCheckboxesChecked('attNames');

  // バッチチェックボックスの状態を更新
  document.getElementById('batchCheck_insiderReci').checked = isInsiderDomainsChecked;
  document.getElementById('batchCheck_outsiderReci').checked = isOutsiderDomainsChecked;
  document.getElementById('batchCheck_Attachments').checked = isAttachmentsChecked;

  // 送信ボタンの有効/無効と色を切り替え
  const sendButton = document.getElementById("btn_send");
  let isAllChecked;
  if (prefs.get("noDisplayInsiderDomainOnly") && outsiderAddressCount === 0) {
    isAllChecked = true; // 組織内ドメインのみの場合は送信ボタンを有効にする
  } else {
    isAllChecked = isInsiderDomainsChecked && isOutsiderDomainsChecked && isMailHeadChecked && isAttachmentsChecked;
  }
  sendButton.disabled = !isAllChecked;
  if (isAllChecked) {
    sendButton.classList.remove('bg-gray-500', 'hover:bg-gray-600', 'dark:bg-gray-600', 'dark:hover:bg-gray-700');
    sendButton.classList.add('bg-blue-500', 'hover:bg-blue-600', 'dark:bg-blue-600', 'dark:hover:bg-blue-700');
  } else {
    sendButton.classList.remove('bg-blue-500', 'hover:bg-blue-600', 'dark:bg-blue-600', 'dark:hover:bg-blue-700');
    sendButton.classList.add('bg-gray-500', 'hover:bg-gray-600', 'dark:bg-gray-600', 'dark:hover:bg-gray-700');
  }
}

function cancelSend() {
  console.log("capopup.js: cancelSend 実行");
  const msgTo = { type: "cancel" };
  Office.context.ui.messageParent(JSON.stringify(msgTo));
}

function confirmSend() {
  console.log("capopup.js: confirmSend実行");
  const prefs = Office.context.roamingSettings;
  if (prefs.get("countDown")) {
    const seconds = parseInt(prefs.get("countDownTime") || 5, 10);
    document.getElementById("countdown").textContent = `あと${seconds}秒で送信します。`;
    Office.context.ui.messageParent(JSON.stringify({ type: "countDown", seconds }));
  } else {
    const msgTo = { type: "confirm" };
    Office.context.ui.messageParent(JSON.stringify(msgTo));
  } 
}