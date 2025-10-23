/******/ (function() { // webpackBootstrap
/*!********************************!*\
  !*** ./src/capopup/capopup.js ***!
  \********************************/
/* global Office, console, document */

console.log("capopup.js: スクリプトロード開始");
var isCountingDown = false;
Office.onReady(function (info) {
  console.log("capopup.js: Office.js 初期化完了:", JSON.stringify(info));
  var setupCheckboxListener = function setupCheckboxListener(checkboxId) {
    var targetId = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
    var checkbox = document.getElementById(checkboxId);
    checkbox.addEventListener("change", function (event) {
      if (targetId) {
        batchCheck(targetId, event.target.checked);
      }
      checkAllChecked();
    });
  };
  setupCheckboxListener("batchCheck_insiderReci", "insiderReci");
  setupCheckboxListener("check_firstLinesOfBody", null);
  var sendButton = document.getElementById("btn_send");
  sendButton.addEventListener("click", confirmSend);
  var cancelButton = document.getElementById("btn_cancel");
  cancelButton.addEventListener("click", cancelSend);
  var prefs = Office.context.roamingSettings;
  var setupBatchCheckButton = function setupBatchCheckButton(buttonId, prefId) {
    if (prefs.get(prefId)) {
      document.getElementById(buttonId).disabled = false;
    }
  };
  setupBatchCheckButton("batchCheck_insiderReci", "insiderDomainBatchCheck");
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onMessageFromParent, onRegisterMessageComplete);
  Office.context.ui.messageParent(JSON.stringify({
    type: "dialogReady"
  }));
  console.log("capopup.js: 準備完了メッセージを送信");
}).catch(function (error) {
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
  var targetDiv = document.getElementById(targetId);
  var inputs = targetDiv.getElementsByTagName("input");
  for (var i = 0; i < inputs.length; i++) {
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
  var prefs = Office.context.roamingSettings;
  if (prefs.get("confirmMailBody")) {
    document.getElementById("mailBodyContainer").hidden = false;
  }
}
function onMessageFromParent(recv) {
  try {
    var message = JSON.parse(recv.message);
    console.log("capopup.js: メッセージ受信:", message);
    if (message.type === "countdownUpdate") {
      document.getElementById("countdown").textContent = "\u3042\u3068".concat(message.seconds, "\u79D2\u3067\u9001\u4FE1\u3057\u307E\u3059\u3002");
      if (!isCountingDown) {
        isCountingDown = true;
        document.getElementById("btn_send").textContent = "今すぐ送信";
      }
      return;
    }
    var emailDetails = message;
    var prefs = Office.context.roamingSettings;
    var outsiderAddr = emailDetails.outsiderReci.length;
    var attachedFiles = emailDetails.attNames.length;
    if (prefs.get("noDisplayInsiderDomainOnly") && outsiderAddr === 0) {
      checkAllChecked(outsiderAddr, attachedFiles);
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

    // emailDetails.body はプレーンテキスト前提で直接設定
    document.getElementById("body").textContent = emailDetails.body || "本文なし";
    document.getElementById("attNames").textContent = "";
    pushToList({
      targetId: "attNames",
      listType: "Attachments",
      pushingList: emailDetails.attNames
    });
    checkAllChecked();
  } catch (error) {
    console.error("capopup.js: メール詳細の解析エラー:", error);
  }
}
function pushToList(args) {
  var targetDiv = document.getElementById(args.targetId);
  var pushTxtNode;
  for (var i = 0; i < args.pushingList.length; i++) {
    switch (args.listType) {
      case "Addresses":
        pushTxtNode = args.pushingList[i].type + args.pushingList[i].address;
        break;
      case "Attachments":
        pushTxtNode = args.pushingList[i].name;
        break;
    }
    var chkbox = document.createElement("input");
    chkbox.setAttribute("type", "checkbox");
    chkbox.setAttribute("id", "checkbox-".concat(args.targetId, "-").concat(i));
    chkbox.addEventListener("change", function () {
      checkAllChecked();
    });
    var label = document.createElement("label");
    label.setAttribute("for", "checkbox-".concat(args.targetId, "-").concat(i));
    label.appendChild(chkbox);
    label.appendChild(document.createTextNode(pushTxtNode));
    targetDiv.appendChild(label);
    targetDiv.appendChild(document.createElement("br"));
  }
}
function checkAllChecked(outsiderAddr, attachedFiles) {
  var areAllCheckboxesChecked = function areAllCheckboxesChecked(containerId) {
    var container = document.getElementById(containerId);
    var checkboxes = container.getElementsByTagName("input");
    return Array.from(checkboxes).every(function (checkbox) {
      return checkbox.checked;
    });
  };
  var prefs = Office.context.roamingSettings;
  var isInsiderDomainsChecked = areAllCheckboxesChecked("insiderReci");
  var isOutsiderDomainsChecked = outsiderAddr == 0 || areAllCheckboxesChecked("outsiderReci");
  var isMailHeadChecked = true;
  if (prefs.get("confirmMailBody")) {
    isMailHeadChecked = document.getElementById("check_firstLinesOfBody").checked;
  }
  var isAttachmentsChecked = attachedFiles == 0 || areAllCheckboxesChecked("attNames");
  document.getElementById("batchCheck_insiderReci").checked = isInsiderDomainsChecked;
  var sendButton = document.getElementById("btn_send");
  var isAllChecked;
  if (prefs.get("noDisplayInsiderDomainOnly") && outsiderAddr === 0) {
    isAllChecked = true;
  } else {
    isAllChecked = isInsiderDomainsChecked && isOutsiderDomainsChecked && isMailHeadChecked && isAttachmentsChecked;
  }
  sendButton.disabled = !isAllChecked;
  if (isAllChecked) {
    sendButton.classList.remove("bg-gray-500", "hover:bg-gray-600", "dark:bg-gray-600", "dark:hover:bg-gray-700");
    sendButton.classList.add("bg-blue-500", "hover:bg-blue-600", "dark:bg-blue-600", "dark:hover:bg-blue-700");
  } else {
    sendButton.classList.remove("bg-blue-500", "hover:bg-blue-600", "dark:bg-blue-600", "dark:hover:bg-blue-700");
    sendButton.classList.add("bg-gray-500", "hover:bg-gray-600", "dark:bg-gray-600", "dark:hover:bg-gray-700");
  }
}
function cancelSend() {
  console.log("capopup.js: cancelSend 実行");
  document.getElementById("btn_send").textContent = "送信";
  isCountingDown = false;
  Office.context.ui.messageParent(JSON.stringify({
    type: "cancel"
  }));
}
function confirmSend() {
  console.log("capopup.js: confirmSend 実行");
  var prefs = Office.context.roamingSettings;
  if (prefs.get("countDown") && !isCountingDown) {
    var seconds = parseInt(prefs.get("countDownTime") || 5, 10);
    document.getElementById("countdown").textContent = "\u3042\u3068".concat(seconds, "\u79D2\u3067\u9001\u4FE1\u3057\u307E\u3059\u3002");
    document.getElementById("btn_send").textContent = "今すぐ送信";
    isCountingDown = true;
    Office.context.ui.messageParent(JSON.stringify({
      type: "countDown",
      seconds: seconds
    }));
  } else {
    document.getElementById("btn_send").textContent = "送信";
    isCountingDown = false;
    Office.context.ui.messageParent(JSON.stringify({
      type: "confirm"
    }));
  }
}
/******/ })()
;
//# sourceMappingURL=capopup.js.map