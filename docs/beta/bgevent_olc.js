/******/ (function() { // webpackBootstrap
/*!************************************!*\
  !*** ./src/bgevent/bgevent_olc.js ***!
  \************************************/
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _regeneratorValues(e) { if (null != e) { var t = e["function" == typeof Symbol && Symbol.iterator || "@@iterator"], r = 0; if (t) return t.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) return { next: function next() { return e && r >= e.length && (e = void 0), { value: e && e[r++], done: !e }; } }; } throw new TypeError(_typeof(e) + " is not iterable"); }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/* global Office, console */
// Outlook Classic版限定 簡易アドレス確認
// 送信前に宛先をダイアログメッセージで確認（SoftBlock使用）
function onMessageSendHandler(_x) {
  return _onMessageSendHandler.apply(this, arguments);
}
function _onMessageSendHandler() {
  _onMessageSendHandler = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(event) {
    var item, recipients, addrList, confirmMsg, _t;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.p = _context.n) {
        case 0:
          console.log("bgevent_olc.js: Outlook Classic用簡易アドレス確認 開始");
          _context.p = 1;
          item = Office.context.mailbox.item;
          _context.n = 2;
          return collectRecipients(item);
        case 2:
          recipients = _context.v;
          if (!(recipients.length === 0)) {
            _context.n = 3;
            break;
          }
          event.completed({
            allowEvent: true
          });
          return _context.a(2);
        case 3:
          addrList = recipients.join("\n");
          confirmMsg = "\u3010\u30A2\u30C9\u30EC\u30B9\u78BA\u8A8D\u3011\n\n\u4EE5\u4E0B\u306E\u30A2\u30C9\u30EC\u30B9\u306B\u9001\u4FE1\u3057\u307E\u3059\u3002\n".concat(addrList, "\n\n\u3053\u306E\u5185\u5BB9\u3067\u9001\u4FE1\u3057\u3066\u3088\u308D\u3057\u3044\u3067\u3059\u304B\uFF1F\n\n\u203BConfirm-Address for Outlook Classic\u304C\u30A4\u30F3\u30B9\u30C8\u30FC\u30EB\u3055\u308C\u3066\u3044\u308B\u5834\u5408\u3001\u7D9A\u3051\u3066\u9001\u4FE1\u78BA\u8A8D\u30C0\u30A4\u30A2\u30ED\u30B0\u304C\u8868\u793A\u3055\u308C\u307E\u3059\u3002");
          event.completed({
            allowEvent: true,
            errorMessage: confirmMsg
          });
          _context.n = 5;
          break;
        case 4:
          _context.p = 4;
          _t = _context.v;
          console.error("bgevent_olc.js: エラー発生:", _t);
          event.completed({
            allowEvent: true
          });
        case 5:
          return _context.a(2);
      }
    }, _callee, null, [[1, 4]]);
  }));
  return _onMessageSendHandler.apply(this, arguments);
}
function collectRecipients(_x2) {
  return _collectRecipients.apply(this, arguments);
} // IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
function _collectRecipients() {
  _collectRecipients = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(item) {
    var recipients, fields, _loop, _i, _fields;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          recipients = [];
          fields = [{
            name: "to",
            label: "To"
          }, {
            name: "cc",
            label: "Cc"
          }, {
            name: "bcc",
            label: "Bcc"
          }];
          _loop = /*#__PURE__*/_regenerator().m(function _loop() {
            var field, _t2;
            return _regenerator().w(function (_context2) {
              while (1) switch (_context2.p = _context2.n) {
                case 0:
                  field = _fields[_i];
                  if (!(item[field.name] && typeof item[field.name].getAsync === "function")) {
                    _context2.n = 4;
                    break;
                  }
                  _context2.p = 1;
                  _context2.n = 2;
                  return new Promise(function (resolve) {
                    item[field.name].getAsync(function (asyncResult) {
                      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value) {
                        var adds = asyncResult.value.filter(function (r) {
                          return r && r.emailAddress;
                        }).map(function (r) {
                          return "".concat(field.label, ": ").concat(r.emailAddress);
                        });
                        recipients.push.apply(recipients, _toConsumableArray(adds));
                      }
                      resolve();
                    });
                  });
                case 2:
                  _context2.n = 4;
                  break;
                case 3:
                  _context2.p = 3;
                  _t2 = _context2.v;
                  console.warn("bgevent_olc.js: ".concat(field.name, " \u53D6\u5F97\u30A8\u30E9\u30FC"), _t2);
                case 4:
                  return _context2.a(2);
              }
            }, _loop, null, [[1, 3]]);
          });
          _i = 0, _fields = fields;
        case 1:
          if (!(_i < _fields.length)) {
            _context3.n = 3;
            break;
          }
          return _context3.d(_regeneratorValues(_loop()), 2);
        case 2:
          _i++;
          _context3.n = 1;
          break;
        case 3:
          console.log("bgevent_olc.js: 収集したアドレス:", recipients);
          return _context3.a(2, recipients);
      }
    }, _callee2);
  }));
  return _collectRecipients.apply(this, arguments);
}
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
/******/ })()
;
//# sourceMappingURL=bgevent_olc.js.map