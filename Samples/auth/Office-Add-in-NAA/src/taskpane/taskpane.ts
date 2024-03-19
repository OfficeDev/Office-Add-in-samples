/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { ssoGetToken } from "./authConfig";
import { getGraphData } from "./msgraph-helper";
import { writeFileNamesToOfficeDocument} from "./document";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
      const accessToken = await ssoGetToken();
      
      const root = '/me/drive/root/children';
      const params = '?$select=name&$top=10';
      const results = await getGraphData(accessToken,root,params);
     // Get item names from the results
          const itemNames = [];
        const oneDriveItems = results["value"];
        for (let item of oneDriveItems) {
          itemNames.push(item["name"]);
        }
        writeFileNamesToOfficeDocument(itemNames);
  } catch (error) {
    console.error(error);
  }
}


