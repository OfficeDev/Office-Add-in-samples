// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation for Outlook on Windows.

function onMessageRecipientsChangedHandler(event) {
  // External subject tag.
  var externalTag = "[External]";

  // Disclaimer text.
  var disclaimer = '<p style = "color:blue"><i>Caution: This email includes external recipients.</i></p>';  

  // Get current recipients.
  var recipients;

  // 1. Get To recipients.
  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
    // Handle success or error.
    if (asyncResult.status !== Office.asyncResultStatus.Succeeded) {
      console.error("Failed to get To recipients. " + JSON.stringify(asyncResult.error));
    }

    recipients = asyncResult.value;
  });

  // 2. Get Cc recipients.
  Office.context.mailbox.item.cc.getAsync(function (asyncResult) {
    // Handle success or error.
    if (asyncResult.status !== Office.asyncResultStatus.Succeeded) {
      console.error("Failed to get Cc recipients. " + JSON.stringify(asyncResult.error));
    }

    recipients += asyncResult.value;
  });

  // 3. Get Bcc recipients.
  Office.context.mailbox.item.bcc.getAsync(function (asyncResult) {
    // Handle success or error.
    if (asyncResult.status !== Office.asyncResultStatus.Succeeded) {
      console.error("Failed to get Bcc recipients. " + JSON.stringify(asyncResult.error));
    }

    recipients += asyncResult.value;
  });

  // TODO: Dynamically determine current organization.

  // Check if any recipients are external.
  var hasExternal = !recipients.contains("contoso.com");

  if (hasExternal) {
    // Ensure "[External]" is prepended to message subject.
    Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.asyncResultStatus.Succeeded) {
        console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));
      }

      var subject = asyncResult.value;
      if (!subject.contains(externalTag)) {
        subject = `${externalTag} ${subject}`;
        Office.context.mailbox.item.subject.setAsync(subject, function (asyncResult) {
          if (asyncResult.status !== Office.asyncResultStatus.Succeeded) {
            console.error("Failed to set subject. " + JSON.stringify(asyncResult.error));
          }
        });
      }
    });

    // Set disclaimer if there are any external recipients.
    Office.context.mailbox.item.body.appendOnSendAsync(
      disclaimer,
      {
        coercionType: Office.CoercionType.Html
      },
      function (asyncResult) {
        console.error("Failed to set disclaimer for appendOnSend. " + JSON.stringify(asyncResult.error));
      }
    );
  } else {
    // Ensure "[External]" is not part of message subject.
    Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.asyncResultStatus.Succeeded) {
        console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));
      }

      var subject = asyncResult.value;
      if (subject.contains(externalTag)) {
        subject = subject.replace(externalTag, "").trimBefore();
        Office.context.mailbox.item.subject.setAsync(subject, function (asyncResult) {
          if (asyncResult.status !== Office.asyncResultStatus.Succeeded) {
            console.error("Failed to set subject. " + JSON.stringify(asyncResult.error));
          }
        });
      }
    });

    // Clear disclaimer if there aren't any external recipients.
    Office.context.mailbox.item.body.appendOnSendAsync(
      null,
      function (asyncResult) {
        console.error("Failed to clear disclaimer for appendOnSend. " + JSON.stringify(asyncResult.error));
      }
    );
  }
}

// Outlook on the web ignores the following.
if (Office.actions) {
  // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
}
