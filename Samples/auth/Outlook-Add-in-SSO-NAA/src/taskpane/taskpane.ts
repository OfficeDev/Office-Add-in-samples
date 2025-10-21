/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console */

import { AccountManager } from "./authConfig";
import { makeGraphRequest } from "./msgraph-helper";

const accountManager = new AccountManager();
const sideloadMsg = document.getElementById("sideload-msg");
const appBody = document.getElementById("app-body");
const getUserDataButton = document.getElementById("getUserData");
const getUserFilesButton = document.getElementById("getUserFiles");
const userName = document.getElementById("userName");
const userEmail = document.getElementById("userEmail");

// Initialize when Office is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    if (getUserDataButton) {
      getUserDataButton.addEventListener("click", getUserData);
    }
    if (getUserFilesButton) {
      getUserFilesButton.addEventListener("click", getUserFiles);
    }
    // Initialize MSAL.
    accountManager.initialize();
  }
});

/**
 * Writes a list of filenames into the email body.
 * @param fileNameList The list of filenames.
 */
async function writeFileNames(fileNameList: string[]) {
  const item = Office.context.mailbox.item;
  let fileNameBody: string = "";
  fileNameList.map((fileName) => (fileNameBody += "<br/>" + fileName));

  if (item) {
    item.body.setAsync(fileNameBody, {
      coercionType: "html",
    });
  }
}

/**
 * Gets the user data such as name and email and displays it
 * in the task pane.
 */
async function getUserData() {
  const userDataElement = document.getElementById("userData");
  // Specify minimum scopes for the token needed.
  const scopes = ["user.read"];

  const response: { displayName: string; mail: string } = await makeGraphRequest(accountManager, scopes, "/me", "");

  if (userDataElement) {
    userDataElement.style.visibility = "visible";
  }
  if (userName) {
    userName.innerText = response.displayName ?? "";
  }
  if (userEmail) {
    userEmail.innerText = response.mail ?? "";
  }
}

/**
 * Gets the first 10 item names (files or folders) from the user's OneDrive.
 * Inserts the item names into the document.
 */
async function getUserFiles() {
  try {
    const names = await getFileNames(10);

    writeFileNames(names);
  } catch (error) {
    console.error(error);
  }
}

/**
 * Gets item names (files or folders) from the user's OneDrive.
 */
async function getFileNames(count = 10) {
  // Specify minimum scopes for the token needed.
  const scopes = ["Files.Read"];
  const response: { value: { name: string }[] } = await makeGraphRequest(
    accountManager, scopes,
    "/me/drive/root/children",
    `?$select=name&$top=${count}`
  );

  const names = response.value.map((item: { name: string }) => item.name);
  return names;
}
