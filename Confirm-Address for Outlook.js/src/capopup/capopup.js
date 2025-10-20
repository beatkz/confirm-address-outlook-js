/* global Office, console, document */

console.log("capopup.js: スクリプトロード開始");

// HTMLタグを除去してプレーンテキストに変換する関数
function stripHtml(html) {
  try {
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = html;
    let text = tempDiv.textContent || tempDiv.innerText || "";
    text = text.replace(/\s+/g, " ").trim();
    return text || "本文なし";
  } catch (error) {
    console.error("capopup.js: HTMLタグ除去エラー:", error);
    return "本文なし";
  }
}

let isCountingDown = false;

Office.onReady((info) => {
  console.log("capopup.js: Office.js 初期化完了:", JSON.stringify(info));

  const setupCheckboxListener = (checkboxId, targetId = null) => {
    const checkbox = document.getElementById(checkboxId);
    checkbox.addEventListener("change", (event) => {
      if (targetId) {
        batchCheck(targetId, event.target.checked);
      }
      checkAllChecked();
    });
  };

  setupCheckboxListener("batchCheck_insiderReci", "insiderReci");
  setupCheckboxListener("batchCheck_Attachments", "attNames");
  setupCheckboxListener("check_firstLinesOfBody", null);

  const sendButton = document.getElementById("btn_send");
  sendButton.addEventListener("click", confirmSend);
  const cancelButton = document.getElementById("btn_cancel");
  cancelButton.addEventListener("click", cancelSend);

  const prefs = Office.context.roamingSettings;
  const setupBatchCheckButton = (buttonId, prefId) => {
    if (prefs.get(prefId)) {
      document.getElementById(buttonId).disabled = false;
    }
  };
  setupBatchCheckButton("batchCheck_insiderReci", "insiderDomainBatchCheck");
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
  if (emailDetails.insiderReci.length > 0) {
    document.getElementById("insiderContainer").hidden = false;
  }
  if (emailDetails.outsiderReci.length > 0) {
    document.getElementById("outsiderContainer").hidden = false;
  }
  if (emailDetails.attNames.length > 0) {
    document.getElementById("attNamesContainer").hidden = false;
  }

  const prefs = Office.context.roamingSettings;
  if (prefs.get("confirmMailBody")) {
    document.getElementById("mailBodyContainer").hidden = false;
  }
}

function onMessageFromParent(recv) {
  try {
    const message = JSON.parse(recv.message);
    console.log("capopup.js: メッセージ受信:", message);
    if (message.type === "countdownUpdate") {
      document.getElementById("countdown").textContent = `あと${message.seconds}秒で送信します。`;
      if (!isCountingDown) {
        isCountingDown = true;
        document.getElementById("btn_send").textContent = "今すぐ送信";
      }
      return;
    }

    const emailDetails = message;
    const prefs = Office.context.roamingSettings;
    const outsiderAddressCount = emailDetails.outsiderReci.length;
    if (prefs.get("noDisplayInsiderDomainOnly") && outsiderAddressCount === 0) {
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
      pushingList: emailDetails.insiderReci,
    });

    document.getElementById("outsiderReci").textContent = "";
    pushToList({
      targetId: "outsiderReci",
      listType: "Addresses",
      pushingList: emailDetails.outsiderReci,
    });

    document.getElementById("body").textContent = stripHtml(emailDetails.body);
    document.getElementById("attNames").textContent = "";
    pushToList({
      targetId: "attNames",
      listType: "Attachments",
      pushingList: emailDetails.attNames,
    });

    checkAllChecked();
  } catch (error) {
    console.error("capopup.js: メール詳細の解析エラー:", error);
  }
}

function pushToList(args) {
  const targetDiv = document.getElementById(args.targetId);
  let pushTxtNode;

  for (let i = 0; i < args.pushingList.length; i++) {
    switch (args.listType) {
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
      checkAllChecked();
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
  const areAllCheckboxesChecked = (containerId) => {
    const container = document.getElementById(containerId);
    const checkboxes = container.getElementsByTagName("input");
    return Array.from(checkboxes).every((checkbox) => checkbox.checked);
  };
  const prefs = Office.context.roamingSettings;

  const isInsiderDomainsChecked = areAllCheckboxesChecked("insiderReci");
  const isOutsiderDomainsChecked =
    outsiderAddressCount === 0 || areAllCheckboxesChecked("outsiderReci");
  let isMailHeadChecked = true;
  if (prefs.get("confirmMailBody")) {
    isMailHeadChecked = document.getElementById("check_firstLinesOfBody").checked;
  }
  const isAttachmentsChecked = areAllCheckboxesChecked("attNames");

  document.getElementById("batchCheck_insiderReci").checked = isInsiderDomainsChecked;
  document.getElementById("batchCheck_Attachments").checked = isAttachmentsChecked;

  const sendButton = document.getElementById("btn_send");
  let isAllChecked;
  if (prefs.get("noDisplayInsiderDomainOnly") && outsiderAddressCount === 0) {
    isAllChecked = true;
  } else {
    isAllChecked =
      isInsiderDomainsChecked &&
      isOutsiderDomainsChecked &&
      isMailHeadChecked &&
      isAttachmentsChecked;
  }
  sendButton.disabled = !isAllChecked;
  if (isAllChecked) {
    sendButton.classList.remove(
      "bg-gray-500",
      "hover:bg-gray-600",
      "dark:bg-gray-600",
      "dark:hover:bg-gray-700"
    );
    sendButton.classList.add(
      "bg-blue-500",
      "hover:bg-blue-600",
      "dark:bg-blue-600",
      "dark:hover:bg-blue-700"
    );
  } else {
    sendButton.classList.remove(
      "bg-blue-500",
      "hover:bg-blue-600",
      "dark:bg-blue-600",
      "dark:hover:bg-blue-700"
    );
    sendButton.classList.add(
      "bg-gray-500",
      "hover:bg-gray-600",
      "dark:bg-gray-600",
      "dark:hover:bg-gray-700"
    );
  }
}

function cancelSend() {
  console.log("capopup.js: cancelSend 実行");
  document.getElementById("btn_send").textContent = "送信";
  isCountingDown = false;
  Office.context.ui.messageParent(JSON.stringify({ type: "cancel" }));
}

function confirmSend() {
  console.log("capopup.js: confirmSend 実行");
  const prefs = Office.context.roamingSettings;
  if (prefs.get("countDown") && !isCountingDown) {
    const seconds = parseInt(prefs.get("countDownTime") || 5, 10);
    document.getElementById("countdown").textContent = `あと${seconds}秒で送信します。`;
    document.getElementById("btn_send").textContent = "今すぐ送信";
    isCountingDown = true;
    Office.context.ui.messageParent(JSON.stringify({ type: "countDown", seconds }));
  } else {
    document.getElementById("btn_send").textContent = "送信";
    isCountingDown = false;
    Office.context.ui.messageParent(JSON.stringify({ type: "confirm" }));
  }
}
