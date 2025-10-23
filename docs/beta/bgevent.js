/******/ (function() { // webpackBootstrap
/*!********************************!*\
  !*** ./src/bgevent/bgevent.js ***!
  \********************************/
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
/* global Office, console */

console.log("bgevent.js: スクリプトロード開始");

// HTMLタグを除去してプレーンテキストに変換する関数（改行を保持）
function stripHtml(html) {
  try {
    // テキストノードと改行タグを再帰的に処理
    var _extractTextWithBreaks = function extractTextWithBreaks(node) {
      var result = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : [];
      var _iterator = _createForOfIteratorHelper(node.childNodes),
        _step;
      try {
        for (_iterator.s(); !(_step = _iterator.n()).done;) {
          var child = _step.value;
          if (child.nodeType === Node.TEXT_NODE) {
            // テキストノードを追加
            var _text = child.textContent.trim();
            if (_text) {
              result.push(_text);
            }
          } else if (child.nodeType === Node.ELEMENT_NODE) {
            // 改行タグ（<br>, <p>, <div>）を検出して改行を追加
            var tagName = child.tagName.toLowerCase();
            if (tagName === "br" || tagName === "p" || tagName === "div") {
              result.push("\n");
            }
            _extractTextWithBreaks(child, result);
          }
        }
      } catch (err) {
        _iterator.e(err);
      } finally {
        _iterator.f();
      }
      return result;
    }; // テキストと改行を抽出
    // HTMLエスケープされた文字列をデコード
    var textarea = document.createElement("textarea");
    textarea.innerHTML = html;
    var decodedHtml = textarea.value;
    console.log("bgevent.js: stripHtml 入力:", decodedHtml);

    // 一時的なdiv要素を作成
    var tempDiv = document.createElement("div");
    tempDiv.innerHTML = decodedHtml;
    var textArray = _extractTextWithBreaks(tempDiv);
    // テキストを結合、連続する改行を保持し、スペースを整理
    var text = textArray.join("\n").replace(/\s+/g, " ").trim();
    console.log("bgevent.js: stripHtml 出力:", text);
    // 空の場合のフォールバック
    return text || "本文なし";
  } catch (error) {
    console.error("bgevent.js: HTMLタグ除去エラー:", error);
    return "本文なし";
  }
}
Office.onReady(function (info) {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
}).catch(function (error) {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});
Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged, function (result) {
  if (result.status === Office.AsyncResultStatus.Failed) {
    console.error("bgevent.js: ItemChangedハンドラ登録エラー:", result.error.message);
  } else {
    console.log("bgevent.js: ItemChangedハンドラ登録成功");
  }
});
var countDownTimer = null;
function onItemChanged() {
  console.log("bgevent.js: onItemChanged発火");
  clearTimeout(countDownTimer);
  Office.context.ui.closeContainer();
}
function checkAddress(args) {
  var _prefs$get;
  var retryCount = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 0;
  console.log("bgevent.js: checkAddress実行");
  if (!Office.context.mailbox.item.to) {
    console.warn("bgevent.js: 宛先が未定義、処理を中断");
    Office.context.ui.closeContainer();
    return;
  }
  var prefs = Office.context.roamingSettings;
  var insiderDomains = ((_prefs$get = prefs.get("insiderDomains")) === null || _prefs$get === void 0 ? void 0 : _prefs$get.split(",").map(function (d) {
    return d.trim().toLowerCase();
  })) || [];
  var maxBodyLines = parseInt(prefs.get("confirmMailBodyLines") || 5, 10);
  var emailDetails = {
    insiderReci: [],
    outsiderReci: [],
    attNames: [],
    body: ""
  };
  var msgCompFields = Office.context.mailbox.item;
  function resolveEmailAddressesAsync(addType, resolve) {
    msgCompFields[addType].getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("bgevent.js: ".concat(addType, " \u53D6\u5F97\u30A8\u30E9\u30FC:"), result.error.message);
        resolve([]);
        return;
      }
      var recipients = result.value.map(function (recipient) {
        return {
          type: addType === "to" ? "To: " : addType === "cc" ? "Cc: " : "Bcc: ",
          address: recipient.emailAddress.toLowerCase()
        };
      });
      resolve(recipients);
    });
  }
  function resolveAttachmentsAsync(resolve) {
    if (!msgCompFields.attachments) {
      resolve([]);
      return;
    }
    var attachments = msgCompFields.attachments.map(function (att) {
      return {
        name: att.name
      };
    });
    resolve(attachments);
  }
  function resolveBodyAsync(resolve) {
    msgCompFields.body.getAsync(Office.CoercionType.Html, function (htmlResult) {
      if (htmlResult.status === Office.AsyncResultStatus.Failed) {
        console.error("bgevent.js: HTML本文取得エラー:", htmlResult.error.message);
        // テキスト形式でリトライ
        msgCompFields.body.getAsync(Office.CoercionType.Text, function (textResult) {
          if (textResult.status === Office.AsyncResultStatus.Failed) {
            console.error("bgevent.js: テキスト本文取得エラー:", textResult.error.message);
            resolve("本文なし");
          } else {
            console.log("bgevent.js: テキスト本文取得成功");
            // テキストメールはそのまま使用
            var body = textResult.value || "本文なし";
            if (prefs.get("confirmMailBody")) {
              body = body.split("\n").slice(0, maxBodyLines).join("\n").trim();
            } else {
              body = "";
            }
            resolve(body);
          }
        });
      } else {
        console.log("bgevent.js: HTML本文取得成功");
        // HTMLメールをプレーンテキストに変換
        var body = stripHtml(htmlResult.value) || "本文なし";
        if (prefs.get("confirmMailBody")) {
          body = body.split("\n").slice(0, maxBodyLines).join("\n").trim();
        } else {
          body = "";
        }
        resolve(body);
      }
    });
  }
  Promise.all([new Promise(function (resolve) {
    return resolveEmailAddressesAsync("to", resolve);
  }), new Promise(function (resolve) {
    return resolveEmailAddressesAsync("cc", resolve);
  }), new Promise(function (resolve) {
    return resolveEmailAddressesAsync("bcc", resolve);
  }), new Promise(function (resolve) {
    return resolveAttachmentsAsync(resolve);
  }), new Promise(function (resolve) {
    return resolveBodyAsync(resolve);
  })]).then(function (_ref) {
    var _ref2 = _slicedToArray(_ref, 5),
      toList = _ref2[0],
      ccList = _ref2[1],
      bccList = _ref2[2],
      attList = _ref2[3],
      body = _ref2[4];
    console.log("bgevent.js: データ取得完了");
    var allReci = [].concat(_toConsumableArray(toList), _toConsumableArray(ccList), _toConsumableArray(bccList));
    allReci.forEach(function (recipient) {
      var _recipient$address$sp;
      var domain = (_recipient$address$sp = recipient.address.split("@")[1]) === null || _recipient$address$sp === void 0 ? void 0 : _recipient$address$sp.toLowerCase();
      if (insiderDomains.includes(domain)) {
        emailDetails.insiderReci.push(recipient);
      } else {
        emailDetails.outsiderReci.push(recipient);
      }
    });
    emailDetails.attNames = attList;
    emailDetails.body = body;
    Office.context.ui.displayDialogAsync("https://localhost:3000/capopup.html", {
      height: 50,
      width: 30,
      displayInIframe: true
    }, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("bgevent.js: ダイアログ表示エラー:", result.error.message);
        Office.context.ui.closeContainer();
        return;
      }
      var dialog = result.value;
      console.log("bgevent.js: ダイアログ表示成功");
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (recv) {
        var message = JSON.parse(recv.message);
        console.log("bgevent.js: ダイアログメッセージ受信:", message);
        if (message.type === "dialogReady") {
          dialog.messageChild(JSON.stringify(emailDetails));
        } else if (message.type === "confirm") {
          dialog.close();
          args.completed({
            allowEvent: true
          });
        } else if (message.type === "countDown") {
          var seconds = message.seconds;
          countDownTimer = setInterval(function () {
            seconds--;
            if (seconds <= 0) {
              clearInterval(countDownTimer);
              dialog.close();
              args.completed({
                allowEvent: true
              });
            } else {
              dialog.messageChild(JSON.stringify({
                type: "countdownUpdate",
                seconds: seconds
              }));
            }
          }, 1000);
        } else if (message.type === "cancel") {
          clearInterval(countDownTimer);
          dialog.close();
          args.completed({
            allowEvent: false
          });
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function (recv) {
        console.error("bgevent.js: ダイアログエラー:", JSON.stringify(recv));
        clearInterval(countDownTimer);
        dialog.close();
        args.completed({
          allowEvent: false
        });
      });
    });
  }).catch(function (error) {
    console.error("bgevent.js: データ取得エラー:", error);
    if (retryCount < 3) {
      console.log("bgevent.js: リトライ", retryCount + 1);
      setTimeout(function () {
        return checkAddress(args, retryCount + 1);
      }, 1000);
    } else {
      console.error("bgevent.js: リトライ上限到達、処理中断");
      args.completed({
        allowEvent: false
      });
    }
  });
}
function onMessageSend(args) {
  console.log("bgevent.js: onMessageSend発火");
  clearTimeout(countDownTimer);
  checkAddress(args);
}
/******/ })()
;
//# sourceMappingURL=bgevent.js.map