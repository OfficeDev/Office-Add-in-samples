/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console */

Office.onReady();

function onMessageRecipientsChangedHandler(event) {
  if (event.changedRecipientFields.to) {
    checkForExternalTo(event);
  } else if (event.changedRecipientFields.cc) {
    checkForExternalCc(event);
  } else if (event.changedRecipientFields.bcc) {
    checkForExternalBcc(event);
  }
}

/**
 * Determines if there are any external recipients in the To field.
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
function checkForExternalTo(event) {
  // Get To recipients.
  Office.context.mailbox.item.to.getAsync(
    { asyncContext: event },
    function (asyncResult) {
      const event = asyncResult.asyncContext;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`Failed to get To recipients: ${asyncResult.error.message}`);
        event.completed();
        return;
      }

      const toRecipients = JSON.stringify(asyncResult.value);
      const keyName = "tagExternalTo";
      if (toRecipients != null
        && toRecipients.length > 0
        && toRecipients.includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        _setSessionData(keyName, true, event);
      } else {
        _setSessionData(keyName, false, event);
      }
    }
  );
}
/**
 * Determines if there are any external recipients in the Cc field.
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
function checkForExternalCc(event) {
  // Get Cc recipients.
  Office.context.mailbox.item.cc.getAsync(
    { asyncContext: event },
    function (asyncResult) {
      const event = asyncResult.asyncContext;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`Failed to get Cc recipients: ${asyncResult.error.message}`);
        event.completed();
        return;
      }
      
      const ccRecipients = JSON.stringify(asyncResult.value);
      const keyName = "tagExternalCc";
      if (ccRecipients != null
          && ccRecipients.length > 0
          && ccRecipients.includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        _setSessionData(keyName, true, event);
      } else {
        _setSessionData(keyName, false, event);
      }
    }
  );
}
/**
 * Determines if there are any external recipients in the Bcc field.
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
function checkForExternalBcc(event) {
  // Get Bcc recipients.
  Office.context.mailbox.item.bcc.getAsync(
    { asyncContext: event },
    function (asyncResult) {
      const event = asyncResult.asyncContext;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`Failed to get Bcc recipients: ${asyncResult.error.message}`);
        event.completed();
        return;
      }

      const bccRecipients = JSON.stringify(asyncResult.value);
      const keyName = "tagExternalBcc";
      if (bccRecipients != null
          && bccRecipients.length > 0
          && bccRecipients.includes(Office.MailboxEnums.RecipientType.ExternalUser)) {
        _setSessionData(keyName, true, event);
      } else {
        _setSessionData(keyName, false, event);
      }
    }
  );
}
/**
 * Sets the value of the specified sessionData key.
 * If value is true, also tag as external, else check entire sessionData property bag.
 * @param {string} key The key or name
 * @param {bool} value The value to assign to the key
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
 function _setSessionData(key, value, event) {
  Office.context.mailbox.item.sessionData.setAsync(
    key,
    value.toString(),
    { asyncContext: event },
    function(asyncResult) {
      const event = asyncResult.asyncContext;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`Failed to set ${key} sessionData to ${value}. Error: ${asyncResult.error.message}`);
        event.completed();
        return;
      }

      console.log(`Set sessionData (${key}) to ${value} successfully.`);
      if (value) {
        _tagExternal(value, event);
      } else {
        _checkForExternal(event);
      }
    }
  );
}
/**
 * Checks the sessionData property bag to determine if any field contains external recipients.
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
function _checkForExternal(event) {
  // Get sessionData to determine if any fields have external recipients.
  Office.context.mailbox.item.sessionData.getAllAsync(
    { asyncContext: event },
    function (asyncResult) {
      const event = asyncResult.asyncContext;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`Failed to get all sessionData: ${asyncResult.error.message}`);
        event.completed();
        return;
      }

      const sessionData = JSON.stringify(asyncResult.value);
      if (sessionData != null
        && sessionData.length > 0
        && sessionData.includes("true")) {
        _tagExternal(true, event);
      } else {
        _tagExternal(false, event);
      }
    }
  );
}
/**
 * If there are any external recipients, prepends the subject of the Outlook item
 * with "[External]" and displays a notification. If there are no external
 * recipients, ensures the tag is not present and clears the notification.
 * @param {bool} hasExternal If there are any external recipients
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
function _tagExternal(hasExternal, event) {
  const externalTag = "[External]";

  Office.context.mailbox.item.subject.getAsync({ asyncContext: event }, function (asyncResult) {
    const event = asyncResult.asyncContext;
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(`Failed to get subject: ${asyncResult.error.message}`);
      event.completed();
      return;
    }

    const currentSubject = asyncResult.value;
    let subject = currentSubject;
    if (hasExternal && !currentSubject.includes(externalTag)) {
      subject = `${externalTag} ${currentSubject}`;
    } else if (!hasExternal && currentSubject.startsWith(externalTag)) {
      subject = currentSubject.replace(externalTag, "").trim();
    }

    if (subject === currentSubject) {
      _updateExternalNotification(hasExternal, event);
      return;
    }

    Office.context.mailbox.item.subject.setAsync(
      subject,
      { asyncContext: event },
      function (asyncResult) {
        const event = asyncResult.asyncContext;
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`Failed to set subject: ${asyncResult.error.message}`);
          event.completed();
          return;
        }

        console.log("Set subject successfully.");
        _updateExternalNotification(hasExternal, event);
      }
    );
  });
}

/**
 * Updates the sender's notification without changing the message body.
 * @param {bool} hasExternal If there are any external recipients
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event
 */
function _updateExternalNotification(hasExternal, event) {
  const notificationKey = "tagExternalRecipients";
  const notificationMessages = Office.context.mailbox.item.notificationMessages;

  if (hasExternal) {
    notificationMessages.replaceAsync(
      notificationKey,
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "This email includes external recipients. Please review all recipients before sending.",
        icon: "Icon.16x16",
        persistent: true,
      },
      { asyncContext: event },
      function (asyncResult) {
        const event = asyncResult.asyncContext;
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(
            `Failed to set external recipient notification: ${asyncResult.error.message}`
          );
        }
        event.completed();
      }
    );
    return;
  }

  notificationMessages.getAllAsync({ asyncContext: event }, function (asyncResult) {
    const event = asyncResult.asyncContext;
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(`Failed to get notifications: ${asyncResult.error.message}`);
      event.completed();
      return;
    }

    // Removing a notification that doesn't exist returns an error.
    if (
      !asyncResult.value.some(function (notification) {
        return notification.key === notificationKey;
      })
    ) {
      event.completed();
      return;
    }

    notificationMessages.removeAsync(
      notificationKey,
      { asyncContext: event },
      function (asyncResult) {
        const event = asyncResult.asyncContext;
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(
            `Failed to clear external recipient notification: ${asyncResult.error.message}`
          );
        }
        event.completed();
      }
    );
  });
}

Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);