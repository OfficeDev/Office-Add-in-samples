/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { makeGraphRequest2 } from "./msgraph-helper";
import "unfetch/polyfill";

/* global console, document, Office, window */

let msalAccountManager;
let gAccessToken = "";
let isInternetExplorer = false;

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
    if (getUserDataButton) {
      getUserDataButton.onclick = getUserData;
    }
    if (getUserFilesButton) {
      getUserFilesButton.onclick = getUserFiles;
    }
    // Check if Trident IE11 webview is in use.
    if (navigator.userAgent.indexOf("Trident") !== -1) {
      isInternetExplorer = true;
    } else {
      // Initialize the MSAL v3 (NAA) library.
      // This will dynamically import the auth config code (MSAL v3 won't load in the IE11 webview.)
      const accountModule = await import("./authConfig");
      msalAccountManager = new accountModule.AccountManager();
      msalAccountManager.initialize();
    }
  }
});

async function writeFileNames(fileNameList: string[]) {
  //  const item = Office.context.mailbox.item;
  console.log("file names are:" + fileNameList);
  let fileNameBody: string = "";
  for (let i = 0; i < fileNameList.length; i++) {
    fileNameBody += "<p>" + fileNameList[i] + "</p>";
  }
  userFiles.innerHTML = fileNameBody;
  console.log(fileNameBody);
  // Office.context.mailbox.item.body.setAsync(fileNameBody, {
  //   coercionType: "html",
  // });
}

/**
 * Gets the user data such as name and email and displays it
 * in the task pane.
 */
async function getUserData() {
  const userDataElement = document.getElementById("userData");
  //const userAccount = await accountManager.ssoGetUserIdentity(["user.read"]);
  const token = await getTokenWithDialogApi(true);
  //const idTokenClaims = userAccount.idTokenClaims as { name?: string; preferred_username?: string };
  //console.log(userAccount.accessToken);
  console.log(token);

  if (userDataElement) {
    userDataElement.style.visibility = "visible";
  }
  // if (userName) {
  //   userName.innerText = idTokenClaims.name ?? "";
  // }
  // if (userEmail) {
  //   userEmail.innerText = idTokenClaims.preferred_username ?? "";
  // }
}

/**
 * Gets the first 10 item names (files or folders) from the user's OneDrive.
 * Inserts the item names into the document.
 */
async function getUserFiles() {
  try {
    console.log("going to get the anmes");
    const names = await getFileNames(10);
    console.log("got hte names" + names);
    writeFileNames(names);
  } catch (error) {
    console.error(error);
  }
}

/**
 * Gets item names (files or folders) from the user's OneDrive.
 */
async function getFileNames(count = 10) {
  try {
    let accessToken = "";
    // Specify minimum scopes for the token needed.
    //const accessToken = await accountManager.ssoGetToken(["Files.Read"]);
    if (gAccessToken !== "") {
      accessToken = gAccessToken;
    } else {
      // accessToken = await getTokenWithDialogApi(true);
      accessToken = await getAccessToken(["Files.Read"]);
      gAccessToken = accessToken;
      console.log(gAccessToken);
      console.log(accessToken);
    }
    let names = [];
    const response: { value: { name: string }[] } = await makeGraphRequest2(
      accessToken,
      "/me/drive/root/children",
      `?$select=name&$top=${count}`
    );
    for (let i = 0; i < response.value.length; i++) {
      names.push(response.value[i].name);
    }
    console.log("names response: " + names);
    return names;
    // makeGraphRequest2(accessToken, "/me/drive/root/children", `?$select=name&$top=${count}`).then((response) => {
    //   console.log(response);
    //   let names: string[] = [];
    //   if ("response.value" + response.value) {
    //     console.log(response.value);
    //     if (response.value.length) {
    //       console.log("lenghth +" + response.value.length);
    //     }
    //   }

    //   for (let i = 0; i < response.value.length; i++) {
    //     names.push(response.value[i]);
    //   }
    //   console.log("names" + names);
    //   return names;
    // });
  } catch (error) {
    console.error("error: " + error);
  }
}

/**
 * Gets an access token for requested scopes.
 */
async function getAccessToken(scopes) {
  // If IE11 webview is in use, call getTokenWithDIalogApi(true) to use the MSAL v2 compatible library.
  if (isInternetExplorer) {
    return getTokenWithDialogApi(true);
  } else {
    // Use the MSAL v3 NAA library.
    return msalAccountManager.ssoGetAccessToken(scopes);
  }
}

async function getTokenWithDialogApi(isInternetExplorer?: boolean): Promise<string> {
  // following code not possible in trident. Is there a way to get auth context in trident?
  //const accountContext = await getAccountContext();
  if (gAccessToken !== "") return gAccessToken;
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      // createLocalUrl(
      //   `${isInternetExplorer ? "dialogie.html" : "dialog.html"}?accountContext=${encodeURIComponent(JSON.stringify(accountContext))}`
      // ),
      createLocalUrl(`${isInternetExplorer ? "dialogie.html" : "dialog.html"}`),
      { height: 60, width: 30 },
      (result) => {
        result.value.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: { message: string; origin: string | undefined }) => {
            const parsedMessage = JSON.parse(arg.message);

            resolve(parsedMessage.token);
            result.value.close();
          }
        );
      }
    );
  });
}

function createLocalUrl(path: string) {
  return `${window.location.origin}/${path}`;
}
