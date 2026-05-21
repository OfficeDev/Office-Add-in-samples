/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(() => {
  // Set up event handler for ItemChanged.
  // This will handle the scenario where the user pins the taskpane open and selects a new message.
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewMessage);

  loadNewMessage();
});

function loadNewMessage() {
  const subject = Office.context.mailbox.item.subject;
  document.getElementById("subject").textContent = subject;

  resetPage();

  // Get the initialization context (if present).
  fetchInitializationContext();

  // Register for InitializationContextChanged.
  Office.context.mailbox.item.addHandlerAsync(
    Office.EventType.InitializationContextChanged,
    loadNewInitContext
  );
}

function resetPage() {
  document.getElementById("error").style.display = "none";
  document.getElementById("error-msg").textContent = "";
  document.getElementById("no-context").style.display = "none";
  document.getElementById("has-context").style.display = "none";
  document.getElementById("init-context").textContent = "";
}

function loadNewInitContext() {
  resetPage();
  fetchInitializationContext();
}

function fetchInitializationContext(retryCount) {
  retryCount = retryCount || 0;

  Office.context.mailbox.item.getInitializationContextAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value) {
        const context = typeof asyncResult.value === "string"
          ? JSON.parse(asyncResult.value)
          : asyncResult.value;
        document.getElementById("init-context").textContent = JSON.stringify(context, null, 2);
        document.getElementById("has-context").style.display = "block";
      } else if (retryCount < 3) {
        setTimeout(function () { fetchInitializationContext(retryCount + 1); }, 1000);
      } else {
        document.getElementById("no-context").style.display = "block";
      }
    } else {
      if (asyncResult.error.code === 9020 && retryCount < 3) {
        setTimeout(function () { fetchInitializationContext(retryCount + 1); }, 1000);
      } else if (asyncResult.error.code === 9020) {
        document.getElementById("no-context").style.display = "block";
      } else {
        showError(JSON.stringify(asyncResult.error, null, 2));
      }
    }
  });
}

function showError(message) {
  document.getElementById("error-msg").textContent = message;
  document.getElementById("error").style.display = "block";
}
