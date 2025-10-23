/******/ (function() { // webpackBootstrap
/*!********************************!*\
  !*** ./src/bgevent/bgevent.js ***!
  \********************************/
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t.return || t.return(); } finally { if (u) throw o; } } }; }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
/* global Office, console, setInterval, clearInterval, process */

// ダイアログを表示
var caDialog; // confirmダイアログのグローバル変数
var countdownInterval = null; // カウントダウン用のグローバル変数

Office.onReady(function (info) {
  console.log("bgevent.js: Office.js 初期化完了:", JSON.stringify(info));
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}).catch(function (error) {
  console.error("bgevent.js: Office.js 初期化エラー:", error);
});

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
    console.log("textArray: ", textArray);
    // テキストを結合、連続する改行を保持し、スペースを整理
    var text = textArray.join("\n").replace(/[ \t]+/g, " ").replace(/\n+/g, "\n").trim();
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
  var dialogUrl = "".concat("https://beatkz.github.io/confirm-address-outlook-js/beta/", "capopup.html");
  console.log("bgevent.js: ダイアログ表示を試行", dialogUrl);
  Office.context.ui.displayDialogAsync(dialogUrl, {
    height: 50,
    width: 30,
    promptBeforeOpen: false
  }, function (result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("bgevent.js: ダイアログ表示エラー:", result.error.message);
      sendEvent.completed({
        allowEvent: false,
        errorMessage: "\u78BA\u8A8D\u753B\u9762\u306E\u8868\u793A\u306B\u5931\u6557: ".concat(result.error.message)
      });
      return;
    }
    console.log("bgevent.js: ダイアログ表示成功");
    caDialog = result.value;
    console.log("bgevent.js: caDialog オブジェクト:", caDialog, "タイプ:", _typeof(caDialog));

    // Office.js のイベントハンドラ
    caDialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
      console.log("bgevent.js: ダイアログからのメッセージ受信:", arg.message);
      handleMessage(arg.message, sendEvent, caDialog);
    });
    caDialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
      console.log("bgevent.js: ダイアログが閉じられました:", arg);
    });
  });
}

// メッセージを処理
function handleMessage(recv, sendEvent, dialog) {
  var msgDlg = JSON.parse(recv);
  switch (msgDlg.type) {
    case "dialogReady":
      console.log("bgevent.js: ダイアログ準備完了メッセージを受信、メール詳細を送信");
      sendEmailDetails();
      break;
    case "confirm":
      console.log("bgevent.js: 確認メッセージを受信、送信を許可");
      dialog.close();
      sendEvent.completed({
        allowEvent: true
      });
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
        errorMessage: "送信がキャンセルされました。"
      });
      break;
    default:
      console.warn("bgevent.js: 無効なメッセージを無視:", recv, "タイプ:", _typeof(recv));
  }
}
function startCountdown(seconds, sendEvent, dialog) {
  var remaining = seconds;
  countdownInterval = setInterval(function () {
    if (remaining <= 0) {
      console.log("bgevent.js: カウントダウン終了、送信を許可");
      clearInterval(countdownInterval);
      countdownInterval = null;
      dialog.close();
      sendEvent.completed({
        allowEvent: true
      });
      return;
    }
    console.log("bgevent.js: カウントダウン残り:", remaining);
    dialog.messageChild(JSON.stringify({
      type: "countdownUpdate",
      seconds: remaining
    }));
    remaining--;
  }, 1000);
}
function checkAddress() {
  return _checkAddress.apply(this, arguments);
}
function _checkAddress() {
  _checkAddress = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
    var msgCompFields, toList, ccList, bccList, domainList, insiderReci, outsiderReci, lines, body, htmlResult, rawBody, hasHtmlTags, attNames;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          console.log("bgevent.js: checkAddress 開始");
          msgCompFields = Office.context.mailbox.item;
          toList = [];
          ccList = [];
          bccList = [];
          _context.n = 1;
          return collectAddress(msgCompFields, toList, ccList, bccList);
        case 1:
          console.log("bgevent.js: メールアドレス収集完了");
          console.log("bgevent.js: To:", toList, "Cc:", ccList, "Bcc:", bccList);
          domainList = getDomainList(); // 組織のドメインリスト
          console.log("bgevent.js: 組織のドメインリスト:", domainList);
          insiderReci = [];
          outsiderReci = [];
          judgeAddress(toList, domainList, insiderReci, outsiderReci);
          judgeAddress(ccList, domainList, insiderReci, outsiderReci);
          judgeAddress(bccList, domainList, insiderReci, outsiderReci);
          console.log("bgevent.js: 組織内アドレス:", insiderReci.map(function (r) {
            return r.address;
          }).join(", "));
          console.log("bgevent.js: 組織外アドレス:", outsiderReci.map(function (r) {
            return r.address;
          }).join(", "));

          // 本文冒頭
          lines = Office.context.roamingSettings.get("confirmMailBodyLines") || 5;
          _context.n = 2;
          return new Promise(function (resolve) {
            return msgCompFields.body.getAsync("html", resolve);
          });
        case 2:
          htmlResult = _context.v;
          console.log("bgevent.js: HTML本文取得完了");
          rawBody = htmlResult.value; // HTMLタグの有無を判定
          hasHtmlTags = /<\/?[a-z][^>]*>/i.test(rawBody);
          console.log("bgevent.js: HTMLタグ検知:", hasHtmlTags);
          if (hasHtmlTags) {
            // HTMLタグが存在する場合、stripHtmlでプレーンテキストに変換
            body = stripHtml(rawBody);
          } else {
            // HTMLタグがない場合（テキストメール）、そのまま使用
            body = rawBody;
          }
          if (!body) {
            body = "本文なし";
          }
          if (Office.context.roamingSettings.get("confirmMailBody")) {
            body = body.split("\n").slice(0, lines).join("\n").trim();
          } else {
            body = "";
          }
          attNames = [];
          _context.n = 3;
          return getAttachments(msgCompFields, attNames);
        case 3:
          console.log("bgevent.js: 添付ファイル名:", attNames.map(function (att) {
            return att.name;
          }).join(", "));
          return _context.a(2, {
            insiderReci: insiderReci,
            outsiderReci: outsiderReci,
            body: body,
            attNames: attNames
          });
      }
    }, _callee);
  }));
  return _checkAddress.apply(this, arguments);
}
function judgeAddress(addressArray, domainList, insiderAddress, outsiderAddress) {
  console.log("bgevent.js: judgeAddress 開始");
  console.log("[JUDGE] " + addressArray.map(function (a) {
    return a.address;
  }).join(", ") + "\n");

  // domainListが空の場合、全て外部とみなす
  if (domainList.length === 0) {
    var _iterator2 = _createForOfIteratorHelper(addressArray),
      _step2;
    try {
      for (_iterator2.s(); !(_step2 = _iterator2.n()).done;) {
        var address = _step2.value;
        outsiderAddress.push(address);
      }
    } catch (err) {
      _iterator2.e(err);
    } finally {
      _iterator2.f();
    }
    return;
  }

  // 登録されたドメインリストとアドレスを比較
  var _iterator3 = _createForOfIteratorHelper(addressArray),
    _step3;
  try {
    for (_iterator3.s(); !(_step3 = _iterator3.n()).done;) {
      var target = _step3.value;
      var _address = target.address;
      if (_address.length === 0) {
        continue;
      }
      var domain = _address.substring(_address.indexOf("@")).toLowerCase();
      var match = false;
      var _iterator4 = _createForOfIteratorHelper(domainList),
        _step4;
      try {
        for (_iterator4.s(); !(_step4 = _iterator4.n()).done;) {
          var insiderDomain = _step4.value;
          if (domain.includes(insiderDomain.toLowerCase())) {
            match = true;
            break;
          }
        }
      } catch (err) {
        _iterator4.e(err);
      } finally {
        _iterator4.f();
      }
      if (match) {
        insiderAddress.push(target);
      } else {
        outsiderAddress.push(target);
      }
    }
  } catch (err) {
    _iterator3.e(err);
  } finally {
    _iterator3.f();
  }
}
function getDomainList() {
  console.log("bgevent.js: getDomainList 開始");
  var domainList = [];
  var settings = Office.context.roamingSettings;
  var insiderDomains = settings.get("insiderDomains");
  if (insiderDomains) {
    var domains = insiderDomains.split(",").map(function (domain) {
      return domain.trim();
    });
    var _iterator5 = _createForOfIteratorHelper(domains),
      _step5;
    try {
      for (_iterator5.s(); !(_step5 = _iterator5.n()).done;) {
        var domain = _step5.value;
        if (domain) {
          domainList.push(domain);
        }
      }
    } catch (err) {
      _iterator5.e(err);
    } finally {
      _iterator5.f();
    }
  }
  return domainList;
}
function collectAddress(_x, _x2, _x3, _x4) {
  return _collectAddress.apply(this, arguments);
}
function _collectAddress() {
  _collectAddress = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(msgCompFields, toList, ccList, bccList) {
    var toMap, tempTo, _iterator6, _step6, reci, ccMap, tempCc, _iterator7, _step7, _reci, bccMap, tempBcc, _iterator8, _step8, _reci2;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          console.log("bgevent.js: collectAddress 開始");

          // To
          _context2.n = 1;
          return new Promise(function (resolve) {
            return msgCompFields.to.getAsync(resolve);
          });
        case 1:
          toMap = _context2.v;
          tempTo = toMap.value.map(function (r) {
            return r.emailAddress;
          }) || [];
          _iterator6 = _createForOfIteratorHelper(tempTo);
          try {
            for (_iterator6.s(); !(_step6 = _iterator6.n()).done;) {
              reci = _step6.value;
              if (reci) {
                toList.push({
                  type: "To: ",
                  address: reci
                });
              }
            }

            // Cc
          } catch (err) {
            _iterator6.e(err);
          } finally {
            _iterator6.f();
          }
          _context2.n = 2;
          return new Promise(function (resolve) {
            return msgCompFields.cc.getAsync(resolve);
          });
        case 2:
          ccMap = _context2.v;
          tempCc = ccMap.value.map(function (r) {
            return r.emailAddress;
          }) || [];
          _iterator7 = _createForOfIteratorHelper(tempCc);
          try {
            for (_iterator7.s(); !(_step7 = _iterator7.n()).done;) {
              _reci = _step7.value;
              if (_reci) {
                ccList.push({
                  type: "Cc: ",
                  address: _reci
                });
              }
            }

            // Bcc
          } catch (err) {
            _iterator7.e(err);
          } finally {
            _iterator7.f();
          }
          _context2.n = 3;
          return new Promise(function (resolve) {
            return msgCompFields.bcc.getAsync(resolve);
          });
        case 3:
          bccMap = _context2.v;
          tempBcc = bccMap.value.map(function (r) {
            return r.emailAddress;
          }) || [];
          _iterator8 = _createForOfIteratorHelper(tempBcc);
          try {
            for (_iterator8.s(); !(_step8 = _iterator8.n()).done;) {
              _reci2 = _step8.value;
              if (_reci2) {
                bccList.push({
                  type: "Bcc: ",
                  address: _reci2
                });
              }
            }
          } catch (err) {
            _iterator8.e(err);
          } finally {
            _iterator8.f();
          }
        case 4:
          return _context2.a(2);
      }
    }, _callee2);
  }));
  return _collectAddress.apply(this, arguments);
}
function getAttachments(_x5, _x6) {
  return _getAttachments.apply(this, arguments);
}
function _getAttachments() {
  _getAttachments = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3(msgCompFields, attList) {
    var attResult, tempAtt, _iterator9, _step9, att;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          _context3.n = 1;
          return new Promise(function (resolve) {
            return msgCompFields.getAttachmentsAsync(resolve);
          });
        case 1:
          attResult = _context3.v;
          tempAtt = attResult.value.map(function (att) {
            return att.name;
          }) || [];
          _iterator9 = _createForOfIteratorHelper(tempAtt);
          try {
            for (_iterator9.s(); !(_step9 = _iterator9.n()).done;) {
              att = _step9.value;
              if (att) {
                attList.push({
                  name: att
                });
              }
            }
          } catch (err) {
            _iterator9.e(err);
          } finally {
            _iterator9.f();
          }
        case 2:
          return _context3.a(2);
      }
    }, _callee3);
  }));
  return _getAttachments.apply(this, arguments);
}
function sendEmailDetails() {
  return _sendEmailDetails.apply(this, arguments);
}
function _sendEmailDetails() {
  _sendEmailDetails = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4() {
    var emailDetails;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.n) {
        case 0:
          console.log("bgevent.js: sendEmailDetails 開始");

          // メールの詳細を収集する処理をここに追加
          _context4.n = 1;
          return checkAddress();
        case 1:
          emailDetails = _context4.v;
          console.log("bgevent.js: 詳細:", emailDetails);
          caDialog.messageChild(JSON.stringify(emailDetails));
        case 2:
          return _context4.a(2);
      }
    }, _callee4);
  }));
  return _sendEmailDetails.apply(this, arguments);
}
/******/ })()
;
//# sourceMappingURL=bgevent.js.map