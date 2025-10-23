/******/ (function() { // webpackBootstrap
/*!**********************************!*\
  !*** ./src/settings/settings.js ***!
  \**********************************/
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
/* global Office, console, document, window, setTimeout */

var isInitialized = false;
var prefsJson = {};
if (window.settingsJsLoaded) {
  console.log("settings.js: 既にロード済み、スキップ");
} else {
  var applyTheme = function applyTheme(isDarkTheme) {
    if (isDarkTheme) {
      document.body.classList.add("dark", "bg-gray-900", "text-gray-100");
      document.body.classList.remove("bg-gray-100", "text-gray-900");
      var rootContainers = document.querySelectorAll(".ms-TaskPane-root, .ms-TaskPane-content");
      rootContainers.forEach(function (container) {
        container.classList.add("bg-gray-900", "text-gray-100");
        container.classList.remove("bg-gray-100", "text-gray-900");
        container.style.backgroundColor = "#111827";
        container.style.color = "#F3F4F6";
      });
      document.querySelectorAll("input, button").forEach(function (el) {
        el.classList.add("bg-gray-800", "border-gray-600", "text-gray-100");
        el.classList.remove("bg-white", "border-gray-300", "text-gray-900");
      });
      console.log("settings.js: ダークテーマ配色を適用");
    } else {
      document.body.classList.add("bg-gray-100", "text-gray-900");
      document.body.classList.remove("dark", "bg-gray-900", "text-gray-100");
      var _rootContainers = document.querySelectorAll(".ms-TaskPane-root, .ms-TaskPane-content");
      _rootContainers.forEach(function (container) {
        container.classList.add("bg-gray-100", "text-gray-900");
        container.classList.remove("bg-gray-900", "text-gray-100");
        container.style.backgroundColor = "#F3F4F6";
        container.style.color = "#111827";
      });
      document.querySelectorAll("input, button").forEach(function (el) {
        el.classList.add("bg-white", "border-gray-300", "text-gray-900");
        el.classList.remove("bg-gray-800", "border-gray-600", "text-gray-100");
      });
      console.log("settings.js: ライトテーマ配色を適用");
    }
  };
  var reloadPrefsJson = function reloadPrefsJson() {
    var prefs = Office.context.roamingSettings;
    prefsJson = {
      insiderDomains: {
        type: "text",
        lval: prefs.get("insiderDomains") || "",
        sval: document.getElementById("insiderDomains").value.trim(),
        validate: function validate(val) {
          return val.match(/^([a-zA-Z0-9.-]+,)*[a-zA-Z0-9.-]+$/);
        }
      },
      noDisplayInsiderDomainOnly: {
        type: "checkbox",
        lval: prefs.get("noDisplayInsiderDomainOnly") || false,
        sval: document.getElementById("noDisplayInsiderDomainOnly").checked
      },
      countDown: {
        type: "checkbox",
        lval: prefs.get("countDown") || false,
        sval: document.getElementById("countDown").checked
      },
      countDownTime: {
        type: "text",
        lval: prefs.get("countDownTime") || 5,
        sval: parseInt(document.getElementById("countDownTime").value, 10),
        validate: function validate(val) {
          return !isNaN(val) && val >= 1 && val <= 60;
        }
      },
      confirmMailBody: {
        type: "checkbox",
        lval: prefs.get("confirmMailBody") || false,
        sval: document.getElementById("confirmMailBody").checked
      },
      confirmMailBodyLines: {
        type: "text",
        lval: prefs.get("confirmMailBodyLines") || 5,
        sval: parseInt(document.getElementById("confirmMailBodyLines").value, 10),
        validate: function validate(val) {
          return !isNaN(val) && val >= 1 && val <= 15;
        }
      },
      insiderDomainBatchCheck: {
        type: "checkbox",
        lval: prefs.get("insiderDomainBatchCheck") || false,
        sval: document.getElementById("insiderDomainBatchCheck").checked
      }
    };
  };
  var loadSettings = function loadSettings() {
    reloadPrefsJson();
    for (var _i = 0, _Object$entries = Object.entries(prefsJson); _i < _Object$entries.length; _i++) {
      var _Object$entries$_i = _slicedToArray(_Object$entries[_i], 2),
        key = _Object$entries$_i[0],
        _Object$entries$_i$ = _Object$entries$_i[1],
        type = _Object$entries$_i$.type,
        lval = _Object$entries$_i$.lval;
      if (type === "text") {
        document.getElementById(key).value = lval;
      } else if (type === "checkbox") {
        document.getElementById(key).checked = lval;
      }
    }
    console.log("settings.js: 設定をロード:", prefsJson);
  };
  var validateSettings = function validateSettings() {
    var allClear = true;
    document.getElementById("domainsErr").classList.add("hidden");
    document.getElementById("countdownErr").classList.add("hidden");
    document.getElementById("bodyLinesErr").classList.add("hidden");
    if (!prefsJson["insiderDomains"].validate(prefsJson["insiderDomains"].sval)) {
      document.getElementById("domainsErr").classList.remove("hidden");
      allClear = false;
    }
    if (prefsJson["countDown"].sval && !prefsJson["countDownTime"].validate(prefsJson["countDownTime"].sval)) {
      document.getElementById("countdownErr").classList.remove("hidden");
      allClear = false;
    }
    if (prefsJson["confirmMailBody"].sval && !prefsJson["confirmMailBodyLines"].validate(prefsJson["confirmMailBodyLines"].sval)) {
      document.getElementById("bodyLinesErr").classList.remove("hidden");
      allClear = false;
    }
    if (!allClear) {
      console.error("settings.js: バリデーションエラー");
      var errorDiv = document.createElement("div");
      errorDiv.id = "errorMessage";
      errorDiv.className = document.body.classList.contains("dark") ? "bg-red-700 text-white p-2 mt-2 rounded" : "bg-red-600 text-white p-2 mt-2 rounded";
      errorDiv.textContent = "入力にエラーがあります。確認してください。";
      document.getElementById("settingsForm").appendChild(errorDiv);
      setTimeout(function () {
        return errorDiv.remove();
      }, 3000);
    }
    return allClear;
  };
  var saveSettings = function saveSettings(event) {
    event.preventDefault();
    reloadPrefsJson();
    var allClear = validateSettings();
    if (allClear) {
      var _saveAsync = function saveAsync(result) {
        var retryCount = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 0;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("settings.js: 設定保存エラー:", result.error.message, "詳細:", JSON.stringify(result.error, null, 2));
          if (retryCount < 3) {
            console.log("settings.js: リトライ", retryCount + 1);
            setTimeout(function () {
              return settings.saveAsync(function (r) {
                return _saveAsync(r, retryCount + 1);
              });
            }, 1000);
          } else {
            var errorDiv = document.createElement("div");
            errorDiv.id = "errorMessage";
            errorDiv.className = document.body.classList.contains("dark") ? "bg-red-700 text-white p-2 mt-2 rounded" : "bg-red-600 text-white p-2 mt-2 rounded";
            errorDiv.textContent = "設定の保存に失敗しました。後で再試行してください。";
            document.getElementById("settingsForm").appendChild(errorDiv);
            setTimeout(function () {
              return errorDiv.remove();
            }, 3000);
          }
        } else {
          console.log("settings.js: 設定を保存:", JSON.stringify(prefsJson));
          var successDiv = document.createElement("div");
          successDiv.id = "successMessage";
          successDiv.className = document.body.classList.contains("dark") ? "bg-green-700 text-white p-2 mt-2 rounded" : "bg-green-600 text-white p-2 mt-2 rounded";
          successDiv.textContent = "設定を保存しました。";
          document.getElementById("settingsForm").appendChild(successDiv);
          setTimeout(function () {
            return successDiv.remove();
          }, 3000);
        }
      };
      var settings = Office.context.roamingSettings;
      for (var _i2 = 0, _Object$entries2 = Object.entries(prefsJson); _i2 < _Object$entries2.length; _i2++) {
        var _Object$entries2$_i = _slicedToArray(_Object$entries2[_i2], 2),
          key = _Object$entries2$_i[0],
          val = _Object$entries2$_i[1];
        settings.set(key, val.sval);
      }
      settings.saveAsync(_saveAsync);
    }
  };
  window.settingsJsLoaded = true;
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook && info.platform === Office.PlatformType.PC) {
      console.warn("settings.js: Outlook Classic (Win32) ではサポートされていません。処理を中断します。");
      document.body.innerHTML = "<div id='platformError'>このアドインはOutlook Classicではサポートされていません。<a href='https://github.com/beatkz/confirm-address-outlook/'>Confirm-Address for Outlook Classicをご利用ください。</a></div>";
      return;
    }
    if (isInitialized) {
      console.log("settings.js: 既に初期化済み、スキップ");
      return;
    }
    isInitialized = true;
    console.log("settings.js: Office.js 初期化完了:", JSON.stringify(info));
    var isDarkTheme = false;
    if (Office.context.officeTheme && Office.context.officeTheme.displayPersona === "Dark") {
      console.log("settings.js: Outlook ダークテーマ検出");
      isDarkTheme = true;
    } else if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) {
      console.log("settings.js: システムダークモード検出");
      isDarkTheme = true;
    }
    applyTheme(isDarkTheme);
    loadSettings();
    document.getElementById("settingsForm").addEventListener("submit", saveSettings);
  }).catch(function (error) {
    console.error("settings.js: Office.js 初期化エラー:", error);
  });
}
/******/ })()
;
//# sourceMappingURL=settings.js.map