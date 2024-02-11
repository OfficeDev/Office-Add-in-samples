/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
 */

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  // Get the Base64-encoded EML format of a reported message.
  Office.context.mailbox.item.getAsFileAsync(
    { asyncContext: event },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(
          `Error encountered during message processing: ${asyncResult.error.message}`
        );
        return;
      }

      // Get the user's responses to the options and text box in the preprocessing dialog.
      const spamReportingEvent = asyncResult.asyncContext;
      const reportedOptions = spamReportingEvent.options;
      const additionalInfo = spamReportingEvent.freeText;

      // Run additional processing operations here.

      /**
       * Signals that the spam-reporting event has completed processing.
       * It then moves the reported message to the Junk Email folder of the mailbox, then
       * shows a post-processing dialog to the user. If an error occurs while the message
       * is being processed, the `onErrorDeleteItem` property determines whether the message
       * will be deleted.
       */
      const event = asyncResult.asyncContext;
      event.completed({
        onErrorDeleteItem: true,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
        showPostProcessingDialog: {
          title: "Contoso Spam Reporting",
          description: "Thank you for reporting this message.",
        },
      });
    }
  );
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart
if (
  Office.context.platform === Office.PlatformType.PC ||
  Office.context.platform == null
) {
  Office.actions.associate("onSpamReport", onSpamReport);
}
