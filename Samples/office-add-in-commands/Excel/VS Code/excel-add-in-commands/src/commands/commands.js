/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */

function getData(event) {
  // Implement your custom code here. The following code is a simple example.
  Office.context.document.setSelectedDataAsync(
    "ExecuteFunction works. Button ID=" + event.source.id,
    function (asyncResult) {
      var error = asyncResult.error;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // Show error message.
      } else {
        // Show success message.
      }
    }
  );

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.getData = getData;
