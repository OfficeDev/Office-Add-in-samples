/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { AccountManager } from "./authConfig";
import { getGraphData } from "./msgraph-helper";
import { writeFileNamesToOfficeDocument } from "./document";

const accountManager = new AccountManager();

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word || info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("getUserData").onclick = getUserData;
    document.getElementById("getUserFiles").onclick = getUserFiles;
    accountManager.initialize();
  }
});

/**
 * Gets the user data such as name and email and displays it
 * in the task pane.
 */
async function getUserData() {
  const userAccount = await accountManager.ssoGetUserIdentity();
  console.log(userAccount);
  document.getElementById("userName").innerText = userAccount.idTokenClaims.name;
  document.getElementById("userEmail").innerText = userAccount.idTokenClaims.email;

}

/**
 * Gets the first 10 item names (files or folders) from the user's OneDrive.
 * Inserts the item names into the document.
 */
async function getUserFiles() {
  try {
    const accessToken = await accountManager.ssoGetToken();

    const root = '/me/drive/root/children';
    const params = '?$select=name&$top=10';
    const results = await getGraphData(accessToken, root, params);
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


