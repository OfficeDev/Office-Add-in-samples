/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady((info) => {
  //if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  //}
});

async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}
