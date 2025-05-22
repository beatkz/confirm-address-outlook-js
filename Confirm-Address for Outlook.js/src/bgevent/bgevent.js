/* global Office, console, document */

Office.onReady((info) => {
  // ここでClassic Outlookを判定
  if (info.host === Office.HostType.Outlook && info.platform === Office.PlatformType.PC) {
    console.warn(
      "bgevent.js: Outlook Classic (Win32) ではサポートされていません。処理を中断します。"
    );
    document.body.innerHTML =
      "<div id='platformError'>このアドインはOutlook Classicではサポートされていません。</div>";
    return;
  }

  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("uniqueMessageSendHandler", uniqueMessageSendHandler);
}).catch((error) => {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

// ダイアログを表示
let caDialog; // confirmダイアログのグローバル変数

function showConfirmDialog(sendEvent) {
  console.log("bgevent.js: ダイアログ表示を試行: https://localhost:3000/confirm.html");

  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/capopup.html",
    { height: 50, width: 30 },
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

// メインのイベントハンドラ
function uniqueMessageSendHandler(event) {
  console.log("bgevent.js: uniqueMessageSendHandler 開始");
  showConfirmDialog(event);
}

async function collectEmailDetails() {
  console.log("bgevent.js: collectEmailDetails 開始");

  const item = Office.context.mailbox.item;
  // 宛先
  const toResult = await new Promise((resolve) => item.to.getAsync(resolve));
  const toReci = toResult.value.map((r) => r.emailAddress).join("\n") || "なし";

  // Cc
  const ccResult = await new Promise((resolve) => item.cc.getAsync(resolve));
  const ccReci = ccResult.value.map((r) => r.emailAddress).join("\n") || "なし";

  // Bcc
  const bccResult = await new Promise((resolve) => item.bcc.getAsync(resolve));
  const bccReci = bccResult.value.map((r) => r.emailAddress).join("\n") || "なし";

  // 本文冒頭
  const bodyResult = await new Promise((resolve) => item.body.getAsync("text", resolve));
  const lines = 10;
  const body = bodyResult.value.split("\n").slice(0, lines).join("\n") || "なし";

  // 添付ファイル名
  const attResult = await new Promise((resolve) => item.getAttachmentsAsync(resolve));
  const attNames = attResult.value.map((att) => att.name).join("\n") || "なし";

  return {
    toReci: toReci,
    ccReci: ccReci,
    bccReci: bccReci,
    body: body,
    attNames: attNames,
  };
}

async function sendEmailDetails() {
  console.log("bgevent.js: sendEmailDetails 開始");

  // メールの詳細を収集する処理をここに追加
  const emailDetails = await collectEmailDetails();

  console.log("bgevent.js: 詳細:", emailDetails);
  caDialog.messageChild(JSON.stringify(emailDetails));
}
