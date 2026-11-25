/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, HTMLButtonElement, Office, window */

Office.onReady(() => {
  for (const [id, message] of [
    ["pick-one", "Picked 1"],
    ["pick-two", "Picked 2"],
  ]) {
    const button = document.getElementById(id);
    if (!(button instanceof HTMLButtonElement)) {
      throw new Error(`The dialog button ${id} is missing.`);
    }
    button.disabled = false;
    button.onclick = () => {
      Office.context.ui.messageParent(message, { targetOrigin: window.location.origin });
    };
  }
  const status = document.getElementById("status");
  if (!status) {
    throw new Error("The dialog status element is missing.");
  }
  status.textContent = "Choose 1 or 2.";
});
