/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, Office */

import { openNumberDialog } from "../dialog/dialog-service.ts";
import type { DialogOutcome } from "../dialog/dialog-service.ts";

function writeResult(outcome: DialogOutcome): Promise<void> {
  return new Promise((resolve, reject) => {
    const completed = (result: Office.AsyncResult<void>) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
      } else {
        resolve();
      }
    };

    if (Office.context.host === Office.HostType.Outlook) {
      const item = Office.context.mailbox.item;
      if (!item) {
        reject(new Error("Open an Outlook message before running the dialog command."));
        return;
      }
      const details: Office.NotificationMessageDetails = outcome.isError
        ? {
            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
            message: outcome.message.slice(0, 150),
          }
        : {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: outcome.message.slice(0, 150),
            icon: "Icon.16x16",
            persistent: true,
          };
      item.notificationMessages.replaceAsync("dialog-result", details, completed);
    } else {
      Office.context.document.setSelectedDataAsync(
        outcome.message,
        { coercionType: Office.CoercionType.Text },
        completed
      );
    }
  });
}

export function showDialog(event: Office.AddinCommands.Event): void {
  openNumberDialog(true, (outcome) => {
    writeResult(outcome)
      .catch((error: Error) => {
        console.error(
          "Unable to display the dialog result. Check the document selection or Outlook item.",
          error
        );
      })
      .finally(() => event.completed());
  });
}

Office.onReady(() => {
  Office.actions.associate("showDialog", showDialog);
});
