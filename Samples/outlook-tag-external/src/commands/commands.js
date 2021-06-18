/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = function (reason) {};

/**
 * Handles the OnMessageRecipientsChanged event.
 * @param event {*} The Office event object
 */
function onMessageRecipientsChangedHandler(event) {
  tagExternal(event);
}

/**
 * Determines if there are any external recipients. If there are, updates the
 * subject of the Outlook message and appends a disclaimer to the message body.
 * @param event {*} The Office event object
 */
function tagExternal(event) {
  console.log("tagExternal method"); //debugging

  // Get To recipients.
  console.log("Get To recipients"); //debugging
  Office.context.mailbox.item.to.getAsync(
    {
      "asyncContext": event
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get To recipients. " + JSON.stringify(asyncResult.error));

        // Call event.completed() after all work is done.
        asyncResult.asyncContext.completed();
        return;
      }

      console.log("To recipients: " + JSON.stringify(asyncResult.value)); //debugging
      var toRecipients = asyncResult.value;
      if (toRecipients != null
          && toRecipients.length > 0
          && JSON.stringify(toRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        console.log("To includes external users"); //debugging

        // Update item if needed since external recipients are included.
        _tagExternal(event, true);

        // Call event.completed() after all work is done.
        asyncResult.asyncContext.completed();
        return;
      }

      // Get Cc recipients.
      console.log("Get Cc recipients"); //debugging
      Office.context.mailbox.item.cc.getAsync(
        {
          "asyncContext": event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to get Cc recipients. " + JSON.stringify(asyncResult.error));
    
            // Call event.completed() after all work is done.
            asyncResult.asyncContext.completed();
            return;
          }
          
          console.log("Cc recipients: " + JSON.stringify(asyncResult.value)); //debugging
          var ccRecipients = asyncResult.value;
          if (ccRecipients != null
              && ccRecipients.length > 0
              && JSON.stringify(ccRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
            console.log("Cc includes external users"); //debugging

            // Update item if needed since external recipients are included.
            _tagExternal(event, true);
            
            // Call event.completed() after all work is done.
            asyncResult.asyncContext.completed();
            return;
          }

          // Get Bcc recipients.
          console.log("Get Bcc recipients"); //debugging
          Office.context.mailbox.item.bcc.getAsync(
            {
              "asyncContext": event
            },
            function (asyncResult) {
              // Handle success or error.
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to get Bcc recipients. " + JSON.stringify(asyncResult.error));
        
                // Call event.completed() after all work is done.
                asyncResult.asyncContext.completed();
                return;
              }
        
              console.log("Bcc recipients: " + JSON.stringify(asyncResult.value)); //debugging
              var bccRecipients = asyncResult.value;
              if (bccRecipients != null
                  && bccRecipients.length > 0
                  && JSON.stringify(bccRecipients).includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
                console.log("Bcc includes external users"); //debugging

                // Update item if needed since external recipients are included.
                _tagExternal(event, true);
          
                // Call event.completed() after all work is done.
                asyncResult.asyncContext.completed();
                return;
              }

              // Update item if needed since external recipients aren't included.
              _tagExternal(event, false);

              // Call event.completed() after all work is done.
              event.completed();
          });
        });
      });
}
/**
 * If there are any external recipients, tags the subject of the Outlook item
 * as "External" and appends a disclaimer to the item body. If there are
 * no external recipients, ensures the tag is not present and clears the disclaimer.
 * @param event {*} The Office event object
 * @param hasExternal {bool} If there are any external recipients
 */
function _tagExternal(event, hasExternal) {
  console.log("_tagExternal method"); //debugging

  // External subject tag.
  const externalTag = "[External]";

  if (hasExternal) {
    console.log("External: Get Subject"); //debugging
    // Ensure "[External]" is prepended to the subject.
    Office.context.mailbox.item.subject.getAsync(
      {
        "asyncContext": event
      },
      function (asyncResult) {
        // Handle success or error.
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));

          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
          return;
        }

        console.log("Current Subject: " + JSON.stringify(asyncResult.value)); //debugging
        var subject = asyncResult.value;
        if (!subject.includes(externalTag)) {
          subject = `${externalTag} ${subject}`;
          console.log("Updated Subject: " + subject); //debugging
          Office.context.mailbox.item.subject.setAsync(
            subject,
            {
              "asyncContext": event
            },
            function (asyncResult) {
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to set Subject. " + JSON.stringify(asyncResult.error));

                // Call event.completed() after all work is done.
                asyncResult.asyncContext.completed();
                return;
              }
          });
        }
    });

    // Set disclaimer as there are external recipients.
    var disclaimer = '<p style = "color:blue"><i>Caution: This email includes external recipients.</i></p>';
    console.log("Set disclaimer"); //debugging
    Office.context.mailbox.item.body.appendOnSendAsync(
      disclaimer,
      {
        "coercionType": Office.CoercionType.Html,
        "asyncContext": event
      },
      function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to set disclaimer for appendOnSend. " + JSON.stringify(asyncResult.error));

          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
          return;
        }
      }
    );
  } else {
    console.log("Internal: Get subject"); //debugging
    // Ensure "[External]" is not part of the subject.
    Office.context.mailbox.item.subject.getAsync(
      {
        "asyncContext": event
      },
      function (asyncResult) {
        // Handle success or error.
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to get subject. " + JSON.stringify(asyncResult.error));

          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
          return;
        }

        console.log("Current subject: " + JSON.stringify(asyncResult.value)); //debugging
        var currentSubject = asyncResult.value;
        if (currentSubject.startsWith(externalTag)) {
          var updatedSubject = currentSubject.replace(externalTag, "");
          console.log("Updated subject: " + updatedSubject); //debugging
          var subject = updatedSubject.trim();
          console.log("Trimmed subject: " + subject); //debugging
          Office.context.mailbox.item.subject.setAsync(
            subject,
            {
              "asyncContext": event
            },
            function (asyncResult) {
              if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to set subject. " + JSON.stringify(asyncResult.error));

                // Call event.completed() after all work is done.
                asyncResult.asyncContext.completed();
                return;
              }
          });
        }
    });

    // Clear disclaimer as there aren't any external recipients.
    console.log("Clear disclaimer"); //debugging
    Office.context.mailbox.item.body.appendOnSendAsync(
      null,
      {
        "asyncContext": event
      },
      function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to clear disclaimer for appendOnSend. " + JSON.stringify(asyncResult.error));

          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
          return;
        }
      }
    );
  }
}

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
