/******/ (function() { // webpackBootstrap
/*!*****************************************************************************************************************************************************************!*\
  !*** ../../../../OneDrive - 株式会社ソリトンシステムズ/Junk/デスクトップ/Favs/Develop/confirm-address/forOutlook.js/Confirm-Address for Outlook.js/src/bgevent/bgevent_classic.js ***!
  \*****************************************************************************************************************************************************************/
/* global Office, console */

/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync("text", {
    asyncContext: event
  }, getBodyCallback);
}
function getBodyCallback(asyncResult) {
  var event = asyncResult.asyncContext;
  var body = "";
  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    body = asyncResult.value;
  } else {
    var message = "本文の取得に失敗しました。";
    console.error(message);
    event.completed({
      allowEvent: false,
      errorMessage: message
    });
    return;
  }
  var matches = hasMatches(body);
  if (matches) {
    Office.context.mailbox.item.getAttachmentsAsync({
      asyncContext: event
    }, getAttachmentsCallback);
  } else {
    event.completed({
      allowEvent: true
    });
  }
}
function hasMatches(body) {
  if (body == null || body == "") {
    return false;
  }
  var arrayOfTerms = ["写真", "文書", "添付"];
  for (var index = 0; index < arrayOfTerms.length; index++) {
    var term = arrayOfTerms[index].trim();
    var regex = RegExp(term, 'i');
    if (regex.test(body)) {
      return true;
    }
  }
  return false;
}
function getAttachmentsCallback(asyncResult) {
  var event = asyncResult.asyncContext;
  if (asyncResult.value.length > 0) {
    for (var i = 0; i < asyncResult.value.length; i++) {
      if (asyncResult.value[i].isInline == false) {
        event.completed({
          allowEvent: true
        });
        return;
      }
    }
    event.completed({
      allowEvent: false,
      errorMessage: "Looks like the body of your message includes an image or an inline file. Attach a copy to the message before sending.",
      // TIP: In addition to the formatted message, it's recommended to also set a
      // plain text message in the errorMessage property for compatibility on
      // older versions of Outlook clients.
      errorMessageMarkdown: "Looks like the body of your message includes an image or an inline file. Attach a copy to the message before sending.\n\n**Tip**: For guidance on how to attach a file, see [Attach files in Outlook](https://www.contoso.com/help/attach-files-in-outlook)."
    });
  } else {
    event.completed({
      allowEvent: false,
      errorMessage: "Looks like you're forgetting to include an attachment.",
      // TIP: In addition to the formatted message, it's recommended to also set a
      // plain text message in the errorMessage property for compatibility on
      // older versions of Outlook clients.
      errorMessageMarkdown: "Looks like you're forgetting to include an attachment.\n\n**Tip**: For guidance on how to attach a file, see [Attach files in Outlook](https://www.contoso.com/help/attach-files-in-outlook)."
    });
  }
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
/******/ })()
;
//# sourceMappingURL=bgevent_classic.js.map