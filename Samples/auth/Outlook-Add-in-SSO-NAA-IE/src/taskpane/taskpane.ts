/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { initializeAuthMethod, getAccessToken } from "./authHelper";
import { makeGraphRequest } from "./msgraph-helper";
import "unfetch/polyfill";

/* global console, document, Office */

const sideloadMsg = document.getElementById("sideload-msg");
const appBody = document.getElementById("app-body");
const getUserDataButton = document.getElementById("getUserData");
const getUserFilesButton = document.getElementById("getUserFiles");
const userName = document.getElementById("userName");
const userEmail = document.getElementById("userEmail");
const userFiles = document.getElementById("userFiles");

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";

    await initializeAuthMethod();

    if (getUserDataButton) {
      getUserDataButton.addEventListener("click", getUserData);
    }
    if (getUserFilesButton) {
      getUserFilesButton.addEventListener("click", getUserFiles);
    }
  }
});

/**
 * Writes the file names to the task pane.
 * @param fileNameList The list of file names.
 */
function writeFileNames(fileNameList: string[]) {
  let fileNameBody: string = "";
  for (let i = 0; i < fileNameList.length; i++) {
    fileNameBody += "<p>" + fileNameList[i] + "</p>";
  }
  if (userFiles) {
    userFiles.innerHTML = fileNameBody;
  }
}

/**
 * click event handler for the Get user files button.
 * Gets list of files from User's OneDrive and writes them to the task pane.
 */
async function getUserFiles() {
  const names = await getFileNamesFromMSGraph();
  if (names) {
    writeFileNames(names);
  }
}

/**
 * Gets the user data such as name and email and displays it
 * in the task pane.
 */
async function getUserData() {
  const userDataElement = document.getElementById("userData");

  try {
    // Specify minimum scopes for the token needed.
    const accessToken = await getAccessToken(["user.read"]);

    const response: { displayName: string; mail: string } = await makeGraphRequest(accessToken, "/me", "");

    if (userDataElement) {
      userDataElement.style.visibility = "visible";
    }
    if (userName) {
      userName.innerText = response.displayName ?? "";
    }
    if (userEmail) {
      userEmail.innerText = response.mail ?? "";
    }
  } catch (ex) {
    console.error(ex);
  }
}

/**
 * Gets item names (files or folders) from the user's OneDrive.
 */
async function getFileNamesFromMSGraph(count = 10): Promise<string[] | undefined> {
  try {
    // Specify minimum scopes for the token needed.
    const accessToken = await getAccessToken(["files.read"]);
    const response: { value: { name: string }[] } = await makeGraphRequest(
      accessToken,
      "/me/drive/root/children",
      `?$select=name&$top=${count}`
    );
    let names = [];
    for (let i = 0; i < response.value.length; i++) {
      names.push(response.value[i].name);
    }
    console.log("names response: " + names);
    return names;
  } catch (error) {
    console.error("error: " + error);
  }
}
