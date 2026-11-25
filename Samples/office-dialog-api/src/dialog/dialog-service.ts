/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, URL, window */

export interface DialogOutcome {
  message: string;
  isError: boolean;
}

export function getDialogSupportError(): string | undefined {
  switch (Office.context.host) {
    case Office.HostType.Outlook:
      return Office.context.requirements.isSetSupported("Mailbox", "1.4")
        ? undefined
        : "This Outlook client requires Mailbox 1.4 to open a dialog.";
    case Office.HostType.Word:
    case Office.HostType.Excel:
    case Office.HostType.PowerPoint:
      return Office.context.requirements.isSetSupported("DialogApi", "1.1")
        ? undefined
        : "This Office client requires DialogApi 1.1 to open a dialog.";
    default:
      return "Open this sample in Word, Excel, PowerPoint, or Outlook.";
  }
}

export function openNumberDialog(
  displayInIframe: boolean,
  report: (outcome: DialogOutcome) => void
): void {
  const supportError = getDialogSupportError();
  if (supportError) {
    report({ message: supportError, isError: true });
    return;
  }

  const url = new URL("dialog.html", window.location.href).href;
  Office.context.ui.displayDialogAsync(
    url,
    { height: 40, width: 30, displayInIframe },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        const messages: Record<number, string> = {
          12004: "The dialog URL must use the same trusted domain as the add-in.",
          12005: "The dialog URL must use HTTPS.",
          12007: "A dialog is already open. Choose a number or close it first.",
          12009: "Permission to open the dialog was declined.",
          12011: "Your browser blocked the dialog. Allow pop-ups for this add-in.",
        };
        report({
          message: messages[result.error.code] || result.error.message,
          isError: true,
        });
        return;
      }

      const dialog = result.value;
      let finished = false;
      const finish = (outcome: DialogOutcome, close: boolean) => {
        if (finished) {
          return;
        }
        finished = true;
        if (close) {
          dialog.close();
        }
        report(outcome);
      };

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
        if (!("message" in args)) {
          finish({ message: `Dialog messaging failed (${args.error}).`, isError: true }, true);
          return;
        }
        if (args.origin && args.origin !== window.location.origin) {
          finish(
            { message: "The dialog sent a message from an unexpected origin.", isError: true },
            true
          );
          return;
        }
        if (args.message !== "Picked 1" && args.message !== "Picked 2") {
          finish(
            { message: "The dialog returned an invalid number selection.", isError: true },
            true
          );
          return;
        }
        finish({ message: args.message, isError: false }, true);
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
        if (!("error" in args)) {
          finish({ message: "The dialog returned an unexpected event.", isError: true }, true);
          return;
        }
        const messages: Record<number, string> = {
          12002: "The dialog page could not be found or loaded.",
          12003: "The dialog navigated to a page that does not use HTTPS.",
          12006: "Dialog closed.",
        };
        finish(
          {
            message: messages[args.error] || `Dialog error (${args.error}).`,
            isError: args.error !== 12006,
          },
          args.error !== 12006
        );
      });
    }
  );
}
