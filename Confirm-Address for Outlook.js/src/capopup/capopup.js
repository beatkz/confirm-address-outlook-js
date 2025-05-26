/* global Office, console, document */

console.log("capopup.js: スクリプトロード開始");

// Office.js の初期化
Office.onReady((info) => {
  // ここでClassic Outlookを判定
  if (info.host === Office.HostType.Outlook && info.platform === Office.PlatformType.PC) {
    console.warn(
      "capopup.js: Outlook Classic (Win32) ではサポートされていません。処理を中断します。"
    );
    document.body.innerHTML =
      "<div id='platformError'>このアドインはOutlook Classicではサポートされていません。</div>";
    return;
  }

  // ここから既存の初期化処理
  console.log("capopup.js: Office.js 初期化完了:", JSON.stringify(info));
  console.log("capopup.js: ホスト:", info.host, "プラットフォーム:", info.platform);
  console.log("capopup.js: Office.context:", Office.context);
  console.log("capopup.js: Office.context.ui:", Office.context.ui);

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

// メッセージを処理
function onMessageFromParent(recv) {
  try {
    const emailDetails = JSON.parse(recv.message);
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
   //console.dir(args);
   var targetDiv = document.getElementById(args.targetId);
   var pushTxtNode;

   for (var i = 0; i < args.pushingList.length; i++) {
       switch(args.listType){
           case "Addresses":
               pushTxtNode = args.pushingList[i].type + args.pushingList[i].address;
               break;
           case "Attachments":
               pushTxtNode = args.pushingList[i].name;
               break;
       }
       var chkbox = document.createElement("input");
       chkbox.setAttribute("type", "checkbox");
       chkbox.setAttribute("id", Math.random());
       chkbox.addEventListener("change", (event) => {
           checkAllChecked();
       });
       var label = document.createElement("label");
       label.appendChild(chkbox);
       label.appendChild(document.createTextNode(pushTxtNode));
       targetDiv.appendChild(label);
       targetDiv.appendChild(document.createElement("br"));
   }
}


function onRegisterMessageComplete(result) {
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.error("capopup.js: addHandlerAsync エラー:", result.error.message);
  } else {
    console.log("capopup.js: DialogParentMessageReceived ハンドラ登録成功");
  }
}

function cancelSend() {
  console.log("capopup.js: cancelSend 実行");
  const msgTo = { type: "cancel" };
  Office.context.ui.messageParent(JSON.stringify(msgTo));
}

function confirmSend() {
  // 送信処理をここに追加
  console.log("capopup.js: confirmSend 実行");
  const msgTo = { type: "confirm" };
  Office.context.ui.messageParent(JSON.stringify(msgTo));
}
