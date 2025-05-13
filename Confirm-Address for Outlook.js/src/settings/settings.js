/* global Office, console, document */

let isInitialized = false;

if (window.settingsJsLoaded) {
  console.log("settings.js: 既にロード済み、スキップ");
} else {
  window.settingsJsLoaded = true;

  Office.onReady((info) => {
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

    loadSettings();
    document.getElementById("settingsForm").addEventListener("submit", saveSettings);
  }).catch((error) => {
    console.error("settings.js: Office.js 初期化エラー:", error);
  });

  function loadSettings() {
    const settings = Office.context.roamingSettings;
    const domains = settings.get("domains") || "";
    const skipSelfDomains = settings.get("skipSelfDomains") || false;
    const countdownEnabled = settings.get("countdownEnabled") || false;
    const countdownSeconds = settings.get("countdownSeconds") || 5;
    const showBodyLinesEnabled = settings.get("showBodyLinesEnabled") || false;
    const bodyLines = settings.get("bodyLines") || 3;
    const selfDomainCheckEnabled = settings.get("selfDomainCheckEnabled") || false;
    const otherDomainCheckEnabled = settings.get("otherDomainCheckEnabled") || false;
    const attachmentCheckEnabled = settings.get("attachmentCheckEnabled") || false;

    document.getElementById("domains").value = domains;
    document.getElementById("skipSelfDomains").checked = skipSelfDomains;
    document.getElementById("countdownEnabled").checked = countdownEnabled;
    document.getElementById("countdownSeconds").value = countdownSeconds;
    document.getElementById("showBodyLinesEnabled").checked = showBodyLinesEnabled;
    document.getElementById("bodyLines").value = bodyLines;
    document.getElementById("selfDomainCheckEnabled").checked = selfDomainCheckEnabled;
    document.getElementById("otherDomainCheckEnabled").checked = otherDomainCheckEnabled;
    document.getElementById("attachmentCheckEnabled").checked = attachmentCheckEnabled;

    console.log("settings.js: 設定をロード:", { domains, skipSelfDomains, countdownEnabled, countdownSeconds, showBodyLinesEnabled, bodyLines, selfDomainCheckEnabled, otherDomainCheckEnabled, attachmentCheckEnabled });
  }

  function saveSettings(event) {
    event.preventDefault();
    const settings = Office.context.roamingSettings;

    const domains = document.getElementById("domains").value.trim();
    const skipSelfDomains = document.getElementById("skipSelfDomains").checked;
    const countdownEnabled = document.getElementById("countdownEnabled").checked;
    const countdownSeconds = parseInt(document.getElementById("countdownSeconds").value, 10);
    const showBodyLinesEnabled = document.getElementById("showBodyLinesEnabled").checked;
    const bodyLines = parseInt(document.getElementById("bodyLines").value, 10);
    const selfDomainCheckEnabled = document.getElementById("selfDomainCheckEnabled").checked;
    const otherDomainCheckEnabled = document.getElementById("otherDomainCheckEnabled").checked;
    const attachmentCheckEnabled = document.getElementById("attachmentCheckEnabled").checked;

    let isValid = true;
    document.getElementById("domainsError").classList.add("hidden");
    document.getElementById("countdownError").classList.add("hidden");
    document.getElementById("bodyLinesError").classList.add("hidden");

    if (domains && !domains.match(/^([a-zA-Z0-9.-]+,)*[a-zA-Z0-9.-]+$/)) {
      document.getElementById("domainsError").classList.remove("hidden");
      isValid = false;
    }
    if (countdownEnabled && (isNaN(countdownSeconds) || countdownSeconds < 1 || countdownSeconds > 60)) {
      document.getElementById("countdownError").classList.remove("hidden");
      isValid = false;
    }
    if (showBodyLinesEnabled && (isNaN(bodyLines) || bodyLines < 1 || bodyLines > 10)) {
      document.getElementById("bodyLinesError").classList.remove("hidden");
      isValid = false;
    }

    if (!isValid) {
      console.error("settings.js: バリデーションエラー");
      const errorDiv = document.createElement("div");
      errorDiv.id = "errorMessage";
      errorDiv.className = document.body.classList.contains("dark")
        ? "bg-red-700 text-white p-2 mt-2 rounded"
        : "bg-red-600 text-white p-2 mt-2 rounded";
      errorDiv.textContent = "入力にエラーがあります。確認してください。";
      document.getElementById("settingsForm").appendChild(errorDiv);
      setTimeout(() => errorDiv.remove(), 3000);
      return;
    }

    settings.set("domains", domains);
    settings.set("skipSelfDomains", skipSelfDomains);
    settings.set("countdownEnabled", countdownEnabled);
    settings.set("countdownSeconds", countdownSeconds);
    settings.set("showBodyLinesEnabled", showBodyLinesEnabled);
    settings.set("bodyLines", bodyLines);
    settings.set("selfDomainCheckEnabled", selfDomainCheckEnabled);
    settings.set("otherDomainCheckEnabled", otherDomainCheckEnabled);
    settings.set("attachmentCheckEnabled", attachmentCheckEnabled);

    function saveAsync(result, retryCount = 0) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("settings.js: 設定保存エラー:", result.error.message, "詳細:", JSON.stringify(result.error, null, 2));
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
        console.log("settings.js: 設定を保存:", { domains, skipSelfDomains, countdownEnabled, countdownSeconds, showBodyLinesEnabled, bodyLines, selfDomainCheckEnabled, otherDomainCheckEnabled, attachmentCheckEnabled });
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