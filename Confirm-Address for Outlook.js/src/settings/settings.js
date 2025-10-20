/* global Office, console, document, window, setTimeout */

let isInitialized = false;
let prefsJson = {};

if (window.settingsJsLoaded) {
  console.log("settings.js: 既にロード済み、スキップ");
} else {
  window.settingsJsLoaded = true;

  Office.onReady((info) => {
    // ここでClassic Outlookを判定
    if (info.host === Office.HostType.Outlook && info.platform === Office.PlatformType.PC) {
      console.warn(
        "settings.js: Outlook Classic (Win32) ではサポートされていません。処理を中断します。"
      );
      document.body.innerHTML =
        "<div id='platformError'>このアドインはOutlook Classicではサポートされていません。<a href='https://github.com/beatkz/confirm-address-outlook/'>Confirm-Address for Outlook Classicをご利用ください。</a></div>";
      return;
    }

    if (isInitialized) {
      console.log("settings.js: 既に初期化済み、スキップ");
      return;
    }
    isInitialized = true;

    console.log("settings.js: Office.js 初期化完了:", JSON.stringify(info));
    // テーマ検出と配色変更
    let isDarkTheme = false;
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
  }).catch((error) => {
    console.error("settings.js: Office.js 初期化エラー:", error);
  });

  function applyTheme(isDarkTheme) {
    // 配色を適用
    if (isDarkTheme) {
      document.body.classList.add("dark", "bg-gray-900", "text-gray-100");
      document.body.classList.remove("bg-gray-100", "text-gray-900");
      const rootContainers = document.querySelectorAll(".ms-TaskPane-root, .ms-TaskPane-content");
      rootContainers.forEach((container) => {
        container.classList.add("bg-gray-900", "text-gray-100");
        container.classList.remove("bg-gray-100", "text-gray-900");
        container.style.backgroundColor = "#111827"; // bg-gray-900
        container.style.color = "#F3F4F6"; // text-gray-100
      });
      document.querySelectorAll("input, button").forEach((el) => {
        el.classList.add("bg-gray-800", "border-gray-600", "text-gray-100");
        el.classList.remove("bg-white", "border-gray-300", "text-gray-900");
      });
      console.log("settings.js: ダークテーマ配色を適用");
    } else {
      document.body.classList.add("bg-gray-100", "text-gray-900");
      document.body.classList.remove("dark", "bg-gray-900", "text-gray-100");
      const rootContainers = document.querySelectorAll(".ms-TaskPane-root, .ms-TaskPane-content");
      rootContainers.forEach((container) => {
        container.classList.add("bg-gray-100", "text-gray-900");
        container.classList.remove("bg-gray-900", "text-gray-100");
        container.style.backgroundColor = "#F3F4F6"; // bg-gray-100
        container.style.color = "#111827"; // text-gray-900
      });
      document.querySelectorAll("input, button").forEach((el) => {
        el.classList.add("bg-white", "border-gray-300", "text-gray-900");
        el.classList.remove("bg-gray-800", "border-gray-600", "text-gray-100");
      });
      console.log("settings.js: ライトテーマ配色を適用");
    }
  }

  function reloadPrefsJson() {
    const prefs = Office.context.roamingSettings;
    prefsJson = {
      insiderDomains: {
        type: "text",
        lval: prefs.get("insiderDomains") || "",
        sval: document.getElementById("insiderDomains").value.trim(),
        validate: function (val) {
          return val.match(/^([a-zA-Z0-9.-]+,)*[a-zA-Z0-9.-]+$/);
        },
      },
      noDisplayInsiderDomainOnly: {
        type: "checkbox",
        lval: prefs.get("noDisplayInsiderDomainOnly") || false,
        sval: document.getElementById("noDisplayInsiderDomainOnly").checked,
      },
      countDown: {
        type: "checkbox",
        lval: prefs.get("countDown") || false,
        sval: document.getElementById("countDown").checked,
      },
      countDownTime: {
        type: "text",
        lval: prefs.get("countDownTime") || 5,
        sval: parseInt(document.getElementById("countDownTime").value, 10),
        validate: function (val) {
          return !isNaN(val) && val >= 1 && val <= 60;
        },
      },
      confirmMailBody: {
        type: "checkbox",
        lval: prefs.get("confirmMailBody") || false,
        sval: document.getElementById("confirmMailBody").checked,
      },
      confirmMailBodyLines: {
        type: "text",
        lval: prefs.get("confirmMailBodyLines") || 5,
        sval: parseInt(document.getElementById("confirmMailBodyLines").value, 10),
        validate: function (val) {
          return !isNaN(val) && val >= 1 && val <= 15;
        },
      },
      insiderDomainBatchCheck: {
        type: "checkbox",
        lval: prefs.get("insiderDomainBatchCheck") || false,
        sval: document.getElementById("insiderDomainBatchCheck").checked,
      },
      attachmentBatchCheck: {
        type: "checkbox",
        lval: prefs.get("attachmentBatchCheck") || false,
        sval: document.getElementById("attachmentBatchCheck").checked,
      },
    };
  }

  function loadSettings() {
    reloadPrefsJson();

    for (const [key, { type, lval }] of Object.entries(prefsJson)) {
      if (type === "text") {
        document.getElementById(key).value = lval;
      } else if (type === "checkbox") {
        document.getElementById(key).checked = lval;
      }
    }

    console.log("settings.js: 設定をロード:", prefsJson);
  }

  function validateSettings() {
    let allClear = true;
    document.getElementById("domainsErr").classList.add("hidden");
    document.getElementById("countdownErr").classList.add("hidden");
    document.getElementById("bodyLinesErr").classList.add("hidden");

    if (!prefsJson["insiderDomains"].validate(prefsJson["insiderDomains"].sval)) {
      document.getElementById("domainsErr").classList.remove("hidden");
      allClear = false;
    }

    if (
      prefsJson["countDown"].val &&
      !prefsJson["countDownTime"].validate(prefsJson["countDownTime"].sval)
    ) {
      document.getElementById("countdownErr").classList.remove("hidden");
      allClear = false;
    }

    if (
      prefsJson["confirmMailBody"].val &&
      !prefsJson["confirmMailBodyLines"].validate(prefsJson["confirmMailBodyLines"].sval)
    ) {
      document.getElementById("bodyLinesErr").classList.remove("hidden");
      allClear = false;
    }

    if (!allClear) {
      console.error("settings.js: バリデーションエラー");
      const errorDiv = document.createElement("div");
      errorDiv.id = "errorMessage";
      errorDiv.className = document.body.classList.contains("dark")
        ? "bg-red-700 text-white p-2 mt-2 rounded"
        : "bg-red-600 text-white p-2 mt-2 rounded";
      errorDiv.textContent = "入力にエラーがあります。確認してください。";
      document.getElementById("settingsForm").appendChild(errorDiv);
      setTimeout(() => errorDiv.remove(), 3000);
    }
    return allClear;
  }

  function saveSettings(event) {
    event.preventDefault();
    reloadPrefsJson();

    const allClear = validateSettings();
    if (allClear) {
      const settings = Office.context.roamingSettings;
      for (const [key, val] of Object.entries(prefsJson)) {
        settings.set(key, val.sval);
      }

      function saveAsync(result, retryCount = 0) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(
            "settings.js: 設定保存エラー:",
            result.error.message,
            "詳細:",
            JSON.stringify(result.error, null, 2)
          );
          if (retryCount < 3) {
            console.log("settings.js: リトライ", retryCount + 1);
            setTimeout(() => settings.saveAsync((r) => saveAsync(r, retryCount + 1)), 1000);
          } else {
            const errorDiv = document.createElement("div");
            errorDiv.id = "errorMessage";
            errorDiv.className = document.body.classList.contains("dark")
              ? "bg-red-700 text-white p-2 mt-2 rounded"
              : "bg-red-600 text-white p-2 mt-2 rounded";
            errorDiv.textContent = "設定の保存に失敗しました。後で再試行してください。";
            document.getElementById("settingsForm").appendChild(errorDiv);
            setTimeout(() => errorDiv.remove(), 3000);
          }
        } else {
          console.log("settings.js: 設定を保存:", JSON.stringify(prefsJson));
          const successDiv = document.createElement("div");
          successDiv.id = "successMessage";
          successDiv.className = document.body.classList.contains("dark")
            ? "bg-green-700 text-white p-2 mt-2 rounded"
            : "bg-green-600 text-white p-2 mt-2 rounded";
          successDiv.textContent = "設定を保存しました。";
          document.getElementById("settingsForm").appendChild(successDiv);
          setTimeout(() => successDiv.remove(), 3000);
        }
      }

      settings.saveAsync(saveAsync);
    }
  }
}
