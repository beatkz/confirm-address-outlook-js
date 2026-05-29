/* global Office, console, document, HTMLElement, HTMLInputElement, HTMLButtonElement */

console.log("capopup.js: スクリプトロード開始");

let isCountingDown: boolean = false;
let outsiderCounts: number = 0;
let attachedFileCounts: number = 0;

Office.onReady((info: { host: Office.HostType }) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("capopup.js: Office.js 初期化完了:", JSON.stringify(info));

    // Theme detection and application (dark base by default to suppress flash; switch to light only if light detected)
    let isDarkTheme: boolean = true;  // ダークベースで開始（フラッシュ抑制）
    if (Office.context.officeTheme) {
      console.log("capopup.js: Office theme detected:", JSON.stringify(Office.context.officeTheme));
      const bgColor = Office.context.officeTheme.bodyBackgroundColor || "";
      if (bgColor && parseInt(bgColor.substring(1, 3), 16) >= 0x80) {  // Light background detected
        console.log("capopup.js: ライトテーマ検出");
        isDarkTheme = false;
      } else {
        console.log("capopup.js: Outlook ダークテーマ検出");
      }
    } else if (window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches) {
      console.log("capopup.js: システムライトモード検出");
      isDarkTheme = false;
    } else {
      console.log("capopup.js: ダークをデフォルト適用");
    }
    addinApplyTheme(isDarkTheme);

    const setupCheckboxListener = (checkboxId: string, targetId: string | null = null) => {
      const checkbox: HTMLInputElement = document.getElementById(checkboxId) as HTMLInputElement;
      checkbox.addEventListener("change", (event: any) => {
        if (targetId) {
          batchCheck(targetId, event.target.checked);
        }
        checkAllChecked();
      });
    };
    setupCheckboxListener("chkSenderAddress", null);
    setupCheckboxListener("batchCheck_insiderReci", "insiderReci");
    setupCheckboxListener("check_firstLinesOfBody", null);
    const sendButton: HTMLButtonElement = document.getElementById("btn_send") as HTMLButtonElement;
    sendButton.addEventListener("click", confirmSend);
    const cancelButton: HTMLButtonElement = document.getElementById(
      "btn_cancel"
    ) as HTMLButtonElement;
    cancelButton.addEventListener("click", cancelSend);

    const prefs = Office.context.roamingSettings;
    const setupBatchCheckButton = (buttonId: string, prefId: string) => {
      if (prefs.get(prefId)) {
        const button = document.getElementById(buttonId) as HTMLButtonElement | null;
        if (button) {
          button.disabled = false;
        }
      }
    };
    setupBatchCheckButton("batchCheck_insiderReci", "insiderDomainBatchCheck");

    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      onMessageFromParent,
      onRegisterMessageComplete
    );
    Office.context.ui.messageParent(JSON.stringify({ type: "dialogReady" }));
    console.log("capopup.js: 準備完了メッセージを送信");
  }
}).catch((error) => {
  console.error("capopup.js: Office.js 初期化エラー:", error);
});

function onRegisterMessageComplete(result: Office.AsyncResult<void>) {
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.error("capopup.js: addHandlerAsync エラー:", result.error.message);
  } else {
    console.log("capopup.js: DialogParentMessageReceived ハンドラ登録成功");
  }
}

function addinApplyTheme(isDarkTheme: boolean) {
  const body = document.body;
  const rootContainers = document.querySelectorAll(
    ".ms-TaskPane-root, .ms-TaskPane-content, .ms-Dialog-content, .ms-Dialog-inner, .ms-Dialog-main"
  );
  if (isDarkTheme) {
    body.classList.add("dark", "is-dark");
    body.classList.remove("light");
    rootContainers.forEach((container) => {
      (container as HTMLElement).classList.add("dark", "is-dark");
      (container as HTMLElement).style.backgroundColor = "#1b1a19";
      (container as HTMLElement).style.color = "#f3f2f1";
    });
    console.log("capopup.js: ダークテーマ配色を適用 (Fabric + custom CSS)");
  } else {
    body.classList.add("light");
    body.classList.remove("dark", "is-dark");
    rootContainers.forEach((container) => {
      (container as HTMLElement).classList.remove("dark", "is-dark");
      (container as HTMLElement).style.backgroundColor = "#ffffff";
      (container as HTMLElement).style.color = "#323130";
    });
    console.log("capopup.js: ライトテーマ配色を適用 (Fabric + custom CSS)");
  }
  // Input/button styles now primarily handled by common.css .dark / .ms-Fabric rules
}

function batchCheck(targetId: string, val: boolean) {
  const targetDiv = document.getElementById(targetId) as HTMLElement | null;
  if (targetDiv) {
    const inputs = targetDiv.getElementsByTagName("input");
    for (let i = 0; i < inputs.length; i++) {
      inputs[i].checked = val;
    }
  }
}

function handleHiddenContainers(emailDetails: any) {
  if (emailDetails.caReciList.insider.length > 0) {
    const insiderContainer = document.getElementById("insiderContainer") as HTMLElement | null;
    if (insiderContainer) {
      insiderContainer.hidden = false;
    }
  }
  if (emailDetails.caReciList.outsider.length > 0) {
    const outsiderContainer = document.getElementById("outsiderContainer") as HTMLElement | null;
    if (outsiderContainer) {
      outsiderContainer.hidden = false;
    }
  }
  if (emailDetails.attNames.length > 0) {
    const attNamesContainer = document.getElementById("attNamesContainer") as HTMLElement | null;
    if (attNamesContainer) {
      attNamesContainer.hidden = false;
    }
  }

  const prefs = Office.context.roamingSettings;
  if (prefs.get("confirmMailBody")) {
    const mailBodyContainer = document.getElementById("mailBodyContainer") as HTMLElement | null;
    if (mailBodyContainer) {
      mailBodyContainer.hidden = false;
    }
  }
}

function onMessageFromParent(recv: { message: string }) {
  try {
    const message = JSON.parse(recv.message);
    const countdownContainer = document.getElementById("countdown") as HTMLElement | null;
    console.log("capopup.js: メッセージ受信:", message);
    if (message.type === "countdownUpdate") {
      if (countdownContainer) {
        countdownContainer.textContent = `あと${message.seconds}秒で送信します。`;
      }
      if (!isCountingDown) {
        isCountingDown = true;
        const sendButton = document.getElementById("btn_send") as HTMLElement | null;
        if (sendButton) {
          sendButton.textContent = "今すぐ送信";
        }
      }
      return;
    }

    const emailDetails = message;
    const prefs = Office.context.roamingSettings;
    outsiderCounts = emailDetails.caReciList.outsider.length;
    attachedFileCounts = emailDetails.attNames.length;

    if (prefs.get("noDisplayInsiderDomainOnly") && outsiderCounts === 0) {
      checkAllChecked();
      confirmSend();
      return;
    }

    handleHiddenContainers(emailDetails);
    console.log("capopup.js: メール詳細を受信:", emailDetails);

    if (emailDetails.senderAddress) {
      const senderAddressDiv = document.getElementById("SenderAddress") as HTMLElement | null;
      if (senderAddressDiv) {
        senderAddressDiv.textContent = emailDetails.senderAddress;
      }
      const senderAddressContainer = document.getElementById(
        "SenderAddressContainer"
      ) as HTMLElement | null;
      if (senderAddressContainer) {
        senderAddressContainer.hidden = false;
      }
    }

    const insiderReciDiv = document.getElementById("insiderReci") as HTMLElement | null;
    if (insiderReciDiv) {
      insiderReciDiv.textContent = "";
    }
    pushToList({
      targetId: "insiderReci",
      listType: "Addresses",
      pushingList: emailDetails.caReciList.insider,
    });

    const outsiderReciDiv = document.getElementById("outsiderReci") as HTMLElement | null;
    if (outsiderReciDiv) {
      outsiderReciDiv.textContent = "";
    }
    pushToList({
      targetId: "outsiderReci",
      listType: "Addresses",
      pushingList: emailDetails.caReciList.outsider,
    });

    // emailDetails.body はプレーンテキスト前提で直接設定
    const bodyDiv = document.getElementById("body") as HTMLElement | null;
    if (bodyDiv) {
      bodyDiv.textContent = emailDetails.body || "本文なし";
    }
    const attNamesDiv = document.getElementById("attNames") as HTMLElement | null;
    if (attNamesDiv) {
      attNamesDiv.textContent = "";
    }
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

function pushToList(args: { targetId: string; listType: string; pushingList: any[] }) {
  const targetDiv = document.getElementById(args.targetId) as HTMLElement | null;
  let pushTxtNode;

  if (!targetDiv) {
    console.error(`capopup.js: 要素が見つかりません: ${args.targetId}`);
    return;
  }

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

function checkAllChecked() {
  const areAllCheckboxesChecked = (containerId: string) => {
    const container = document.getElementById(containerId) as HTMLElement | null;
    if (!container) return false;

    const checkboxes = container.getElementsByTagName("input");
    return Array.from(checkboxes).every((checkbox) => checkbox.checked);
  };
  const prefs = Office.context.roamingSettings;

  const chkSenderAddress = document.getElementById("chkSenderAddress") as HTMLInputElement;
  const isSenderAddressChecked = chkSenderAddress.checked;
  const isInsiderDomainsChecked = areAllCheckboxesChecked("insiderReci");
  const isOutsiderDomainsChecked = outsiderCounts == 0 || areAllCheckboxesChecked("outsiderReci");
  let isMailHeadChecked = true;
  if (prefs.get("confirmMailBody")) {
    const chkFirstLines = document.getElementById("check_firstLinesOfBody") as HTMLInputElement;
    isMailHeadChecked = chkFirstLines.checked;
  }
  const isAttachmentsChecked = attachedFileCounts == 0 || areAllCheckboxesChecked("attNames");

  const batchCheckInsider = document.getElementById("batchCheck_insiderReci") as HTMLInputElement;
  batchCheckInsider.checked = isInsiderDomainsChecked;

  const sendButton = document.getElementById("btn_send") as HTMLButtonElement;
  let isAllChecked;
  if (prefs.get("noDisplayInsiderDomainOnly") && outsiderCounts === 0) {
    isAllChecked = true;
  } else {
    isAllChecked =
      isSenderAddressChecked &&
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
  const btnSend = document.getElementById("btn_send") as HTMLButtonElement;
  btnSend.textContent = "送信";
  isCountingDown = false;
  Office.context.ui.messageParent(JSON.stringify({ type: "cancel" }));
}

function confirmSend() {
  console.log("capopup.js: confirmSend 実行");
  const prefs = Office.context.roamingSettings;
  const countdownContainer: HTMLElement | null = document.getElementById("countdown");
  const btnSend = document.getElementById("btn_send") as HTMLButtonElement;
  if (prefs.get("countDown") && !isCountingDown) {
    const seconds = parseInt(prefs.get("countDownTime") || 5, 10);
    if (countdownContainer) {
      countdownContainer.textContent = `あと${seconds}秒で送信します。`;
    }
    btnSend.textContent = "今すぐ送信";
    isCountingDown = true;
    Office.context.ui.messageParent(JSON.stringify({ type: "countDown", seconds }));
  } else {
    btnSend.textContent = "送信";
    isCountingDown = false;
    Office.context.ui.messageParent(JSON.stringify({ type: "confirm" }));
  }
}
