/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import ssoGetToken from "./authConfig";
import getGraphData from "../msgraph-helpers/msgraph-helper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const accessToken = await ssoGetToken();
      const root = '/me/drive/root/children';
      const params = '?$select=name&$top=10';
      const data = await getGraphData(accessToken,root,params);
      // MS Graph data includes OData metadata and eTags that we don't need.
        // Send only what is actually needed to the client: the item names.
        const itemNames = [];
        const oneDriveItems = data["value"];
        for (let item of oneDriveItems) {
          itemNames.push(item["name"]);
        }
        writeFileNamesToWorksheet(itemNames);
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



async function writeFileNamesToWorksheet(result) {
  return Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    var filenames = [];
    var i;
    for (i = 0; i < result.length; i++) {
      var innerArray = [];
      innerArray.push(result[i]);
      filenames.push(innerArray);
    }

    var rangeAddress = 'B5:B' + (5 + (result.length - 1)).toString();
    var range = sheet.getRange(rangeAddress);
    range.values = filenames;
    range.format.autofitColumns();

    return context.sync();
  });
}