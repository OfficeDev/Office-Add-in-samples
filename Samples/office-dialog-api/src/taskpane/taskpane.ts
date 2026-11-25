/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, HTMLButtonElement, Office */

import { getDialogSupportError, openNumberDialog } from "../dialog/dialog-service.ts";
import type { DialogOutcome } from "../dialog/dialog-service.ts";

function showNotification(outcome: DialogOutcome): void {
  const status = document.getElementById("status");
  if (!status) {
    throw new Error("The task pane status element is missing.");
  }
  status.textContent = outcome.message;
  status.classList.toggle("error", outcome.isError);
}

Office.onReady(() => {
  const supportError = getDialogSupportError();
  if (supportError) {
    showNotification({ message: supportError, isError: true });
    return;
  }

  for (const [id, displayInIframe] of [
    ["open-popup", false],
    ["open-iframe", true],
  ] as const) {
    const button = document.getElementById(id);
    if (!(button instanceof HTMLButtonElement)) {
      throw new Error(`The task pane button ${id} is missing.`);
    }
    button.disabled = false;
    button.onclick = () => openNumberDialog(displayInIframe, showNotification);
  }
  showNotification({ message: "Choose how to open the number picker.", isError: false });
});
