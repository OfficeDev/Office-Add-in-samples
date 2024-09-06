/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // showedFRE is created and set to "true" when you call showFirstRunExperience().
    if (!localStorage.getItem("showedFRE")) {
      showFirstRunExperience();
    }

    document.getElementById("run").onclick = run;
  }
});

async function showFirstRunExperience() {
  document.getElementById("first-run-experience").style.display = "flex";
  localStorage.setItem("showedFRE", true);
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
