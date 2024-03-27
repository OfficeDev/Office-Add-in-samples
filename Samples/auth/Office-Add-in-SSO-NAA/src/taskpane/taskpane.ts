/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */

import { AccountManager } from "./authConfig";
import { getGraphData } from "./msgraph-helper";
import { writeFileNamesToOfficeDocument } from "./document";

const accountManager = new AccountManager();
const sideloadMsg = document.getElementById("sideload-msg");
const appBody = document.getElementById("app-body");
const getUserDataButton = document.getElementById("getUserData");
const getUserFilesButton = document.getElementById("getUserFiles");
const userName = document.getElementById("userName");
const userEmail = document.getElementById("userEmail");

Office.onReady((info) => {
  switch (info.host) {
    case Office.HostType.Excel:
    case Office.HostType.PowerPoint:
    case Office.HostType.Word:
      if (sideloadMsg) {
        sideloadMsg.style.display = "none";
      }
      if (appBody) {
        appBody.style.display = "flex";
      }
      if (getUserDataButton) {
        getUserDataButton.onclick = getUserData;
      }
      if (getUserFilesButton) {
        getUserFilesButton.onclick = getUserFiles;
      }
      accountManager.initialize();
      break;
  }
});

/**
 * Gets the user data such as name and email and displays it
 * in the task pane.
 */
async function getUserData() {
  const userDataElement = document.getElementById("userData");
  const userAccount = await accountManager.ssoGetUserIdentity();
  const idTokenClaims = userAccount.idTokenClaims as { name?: string; email?: string };

  console.log(userAccount);

  if (userDataElement) {
    userDataElement.style.visibility = "visible";
  }
  if (userName) {
    userName.innerText = idTokenClaims.name ?? "";
  }
  if (userEmail) {
    userEmail.innerText = idTokenClaims.email ?? "";
  }
}

/**
 * Gets the first 10 item names (files or folders) from the user's OneDrive.
 * Inserts the item names into the document.
 */
async function getUserFiles() {
  try {
    const accessToken = await accountManager.ssoGetToken();

    const root = "/me/drive/root/children";
    const params = "?$select=name&$top=10";
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
