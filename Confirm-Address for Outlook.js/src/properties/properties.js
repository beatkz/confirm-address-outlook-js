/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    run();
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  const properties = document.getElementById("properties");
  properties.innerHTML = "";

  // 宛先
  const toResult = await new Promise((resolve) => item.to.getAsync(resolve));
  const toReci = toResult.value.map((r) => r.emailAddress).join("<br/>") || "なし";
  addProperty(properties, "送信先 [To]", toReci);

  // Cc
  const ccResult = await new Promise((resolve) => item.cc.getAsync(resolve));
  const ccReci = ccResult.value.map((r) => r.emailAddress).join("<br/>") || "なし";
  addProperty(properties, "送信先 [Cc]", ccReci);

  // Bcc
  const bccResult = await new Promise((resolve) => item.bcc.getAsync(resolve));
  const bccReci = bccResult.value.map((r) => r.emailAddress).join("<br/>") || "なし";
  addProperty(properties, "送信先 [Bcc]", bccReci);

  // 本文冒頭
  const bodyResult = await new Promise((resolve) => item.body.getAsync("text", resolve));
  const lines = 8;
  const body = bodyResult.value.split("\n").slice(0, lines).join("<br/>") || "なし";
  addProperty(properties, `本文冒頭 ${lines}行`, body);

  // 添付ファイル名
  const attResult = await new Promise((resolve) => item.getAttachmentsAsync(resolve));
  const attNames = attResult.value.map((att) => att.name).join("<br/>") || "なし";
  addProperty(properties, "添付ファイル名", attNames);
}

function addProperty(list, label, value) {
  const li = document.createElement("li");
  li.innerHTML = `<strong>${label}:</strong><br/>${value}`;
  list.appendChild(li);
}
