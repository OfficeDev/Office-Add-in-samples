/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Initialize the Office JavaScript API library.
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
});

// Configure the keyboard shortcut to open the task pane.
Office.actions.associate("ShowTaskpane", () => {
return Office.addin
  .showAsTaskpane()
  .then(() => {
    return;
  })
  .catch((error) => {
    return error.code;
  });
});

// Configure the keyboard shortcut to close the task pane.
Office.actions.associate("HideTaskpane", () => {
return Office.addin
  .hide()
  .then(() => {
    return;
  })
  .catch((error) => {
    return error.code;
  });
});

// Configure the keyboard shortcut to run an action that's specific to the current Office host.
Office.actions.associate("RunAction", () => {
const host = Office.context.host;

// Cycle through cell colors in Excel.
if (host === Office.HostType.Excel) {
  const context = new Excel.RequestContext();
  const range = context.workbook.getSelectedRange();
  const rangeFormat = range.format;
  rangeFormat.fill.load();
  const colors = ["#FFFFFF", "#C7CC7A", "#7560BA", "#9DD9D2", "#FFE1A8", "#E26D5C"];
  return context.sync().then(() => {
    const rangeTarget = context.workbook.getSelectedRange();
    let currentColor = -1;
    for (let i = 0; i < colors.length; i++) {
      if (colors[i] == rangeFormat.fill.color) {
        currentColor = i;
        break;
      }
    }
    if (currentColor == -1) {
      currentColor = 0;
    } else if (currentColor == colors.length - 1) {
      currentColor = 0;
    } else {
      currentColor++;
    }
    rangeTarget.format.fill.color = colors[currentColor];
    return context.sync();
  });
} else if (host === Office.HostType.Word) {
  // Insert text into the Word document.
  const context = new Word.RequestContext();
  return context.sync().then(() => {
    context.document.body.insertText(
      "Added using a custom keyboard shortcut.",
      Word.InsertLocation.start
    );
    return context.sync();
  });
}
});

// Display the shortcut conflict dialog for testing.
Office.actions.associate("TestConflict", () => {
console.log("Display the shortcut conflict dialog for testing.");
});
