/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
*/

/* global document, Office */

Office.onReady((info) => {
  console.log(`Office.onReady(): Host: ${Office.HostType.Outlook}`);
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";     
  }
});
