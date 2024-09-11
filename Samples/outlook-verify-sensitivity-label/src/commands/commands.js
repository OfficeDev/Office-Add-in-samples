/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Ensures the Office.js library is loaded.
 */
Office.onReady();

/**
 * The legal hold email account of the fictitious company, Fabrikam. It's added to the Bcc field of a
 * message that's configured with the Highly Confidential sensitivity label.
 * @constant
 * @type {string}
 */
const LEGAL_HOLD_ACCOUNT = "legalhold@fabrikam.com";

/**
 * The email address suffix that identifies an account owned by a legal team member at Fabrikam.
 * @constant
 * @type {string}
 */
const LEGAL_TEAM_ACCOUNT_SUFFIX = "-legal@fabrikam.com";

/**
 * Handle the OnMessageRecipientsChanged event by checking whether the legal hold email account was added to the
 * To, Cc, or Bcc field of a message. If the account was added to the To or Cc field, the account is removed.
 * If the account was added to the Bcc field, it's only removed if the sensitivity label of a message isn't set to
 * Highly Confidential.
 * @param {Office.AddinCommands.Event} event The OnMessageRecipientsChanged event object.
 */
function onMessageRecipientsChangedHandler(event) {
  if (event.changedRecipientFields.to) {
    removeLegalHoldAccount(event, Office.context.mailbox.item.to);
  }

  if (event.changedRecipientFields.cc) {
    removeLegalHoldAccount(event, Office.context.mailbox.item.cc);
  }

  if (event.changedRecipientFields.bcc) {
    checkForLegalHoldAccount(event);
  }
}

/**
 * Handle the OnMessageSend event by checking whether the current message has an attachment or a recipient is a member
 * of the legal team. If either of these conditions is true, the event handler checks for the Highly Confidential sensitivity
 * label on the message and sets it if needed.
 * @param {Office.AddinCommands.Event} event The OnMessageSend event object. 
 */
function onMessageSendHandler(event) {
  Office.context.mailbox.item.getAttachmentsAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to check for attachments.");
      console.log(`Error: ${result.error.message}`);
      event.completed({ allowEvent: false, errorMessage: "Unable to check the message for attachments. Save your message, then restart Outlook." });
      return;
    }

    const attachments = result.value;
    if (attachments.length > 0 ) {
      ensureHighlyConfidentialLabelSet(event);
    } else {
      Office.context.mailbox.item.to.getAsync({ asyncContext: event }, (result) => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to get the recipients from the To field.");
          console.log(`Error: ${result.error.message}`);
          event.completed({ allowEvent: false, errorMessage: "Unable to get the recipients from the To field. Save your message, then restart Outlook." });
          return;
        }

        if (containsLegalTeamMember(result.value)) {
          ensureHighlyConfidentialLabelSet(event);
        } else {
          Office.context.mailbox.item.bcc.getAsync({ asyncContext: event }, (result) => {
            const event = result.asyncContext;
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.log("Unable to get the recipients from the Bcc field.");
              console.log(`Error: ${result.error.message}`);
              event.completed({ allowEvent: false, errorMessage: "Unable to get the recipients from the Bcc field. Save your message, then restart Outlook." });
              return;
            }

            if (containsLegalTeamMember(result.value)) {
              ensureHighlyConfidentialLabelSet(event);
            } else {
              Office.context.mailbox.item.cc.getAsync({ asyncContext: event }, (result) => {
                const event = result.asyncContext;
                if (result.status === Office.AsyncResultStatus.Failed) {
                  console.log("Unable to get the recipients from the Cc field.");
                  console.log(`Error: ${result.error.message}`);
                  event.completed({ allowEvent: false, errorMessage: "Unable to get the recipients from the Cc field. Save your message, then restart Outlook." });
                  return;
                }

                if (containsLegalTeamMember(result.value)) {
                  ensureHighlyConfidentialLabelSet(event);
                } else {
                  event.completed({ allowEvent: true });
                }
              });
            }
          });
        }
      });
    }
  });
}

/**
 * Handle the OnSensitivityLabelChanged event by verifying that the legal hold email account is added to the
 * Bcc field if the sensitivity label of the message is set to Highly Confidential. If the sensitivity label
 * isn't set to Highly Confidential, the event handler removes the legal hold email account from the message,
 * if it's present.
 * @param {Office.AddinCommands.Event} event The OnSensitivityLabelChanged event object.
 */
function onSensitivityLabelChangedHandler(event) {
  checkForLegalHoldAccount(event);
}

/**
 * Check that the Highly Confidential sensitivity label is set if a message contains an attachment or a recipient
 * who's a member of the legal team.
 * @param {Office.AddinCommands.Event} event The OnMessageSend event object.
 */
function ensureHighlyConfidentialLabelSet(event) {
  Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve the status of the sensitivity label catalog.");
      console.log(`Error: ${result.error.message}`);
      event.completed({ allowEvent: false, errorMessage: "Unable to retrieve the status of the sensitivity label catalog. Save your message, then restart Outlook." });
      return;
    }

    Office.context.sensitivityLabelsCatalog.getAsync({ asyncContext: event }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to retrieve the catalog of sensitivity labels.");
        console.log(`Error: ${result.error.message}`);
        event.completed({ allowEvent: false, errorMessage: "Unable to retrieve the catalog of sensitivity labels. Save your message, then restart Outlook." });
        return;
      }

      const highlyConfidentialLabel = getLabelId("Highly Confidential", result.value);
      Office.context.mailbox.item.sensitivityLabel.getAsync({ asyncContext: { event: event, highlyConfidentialLabel: highlyConfidentialLabel } }, (result) => {
        const event = result.asyncContext.event;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to get the sensitivity label of the message.");
          console.log(`Error: ${result.error.message}`);
          event.completed({ allowEvent: false, errorMessage: "Unable to get the sensitivity label applied to the message. Save your message, then restart Outlook." });
          return;
        }

        const highlyConfidentialLabel = result.asyncContext.highlyConfidentialLabel;
        if (result.value === highlyConfidentialLabel) {
          event.completed({ allowEvent: true });
        } else {
          Office.context.mailbox.item.sensitivityLabel.setAsync(highlyConfidentialLabel, { asyncContext: event }, (result) => {
            const event = result.asyncContext;
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.log("Unable to set the Highly Confidential sensitivity label to the message.");
              console.log(`Error: ${result.error.message}`);
              event.completed({ allowEvent: false, errorMessage: "Unable to set the Highly Confidential sensitivity label to the message. Save your message, then restart Outlook." });
              return;
            }

            event.completed({ allowEvent: false, errorMessage: "Due to the contents of your message, the sensitivity label has been set to Highly Confidential and the Legal Hold account has been added to the Bcc field.\nTo learn more, see Fabrikam's information protection policy.\n\nDo you need to make changes to your message?" });
          });
        }
      });
    });
  });
}

/**
 * Check whether the legal hold account was added to the Bcc field if the sensitivity label of a message is set to
 * Highly Confidential. If the account appears in the Bcc field, but the sensitivity label isn't set to
 * Highly Confidential, the account is removed from the message.
 * @param {Office.AddinCommands.Event} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
 */
function checkForLegalHoldAccount(event) {
  Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve the status of the sensitivity label catalog.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    Office.context.sensitivityLabelsCatalog.getAsync({ asyncContext: event }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to retrieve the catalog of sensitivity labels.");
        console.log(`Error: ${result.error.message}`);
        event.completed();
        return;
      }

      const highlyConfidentialLabel = getLabelId("Highly Confidential", result.value);
      Office.context.mailbox.item.sensitivityLabel.getAsync({ asyncContext: { event: event, highlyConfidentialLabel: highlyConfidentialLabel, } }, (result) => {
        const event = result.asyncContext.event;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to get the sensitivity label of the message.");
          console.log(`Error: ${result.error.message}`);
          event.completed();
          return;
        }

        if (result.value === result.asyncContext.highlyConfidentialLabel) {
          addLegalHoldAccount(event, Office.context.mailbox.item.bcc);
        } else {
          removeLegalHoldAccount(event, Office.context.mailbox.item.bcc);
        }
      });
    });
  });
}

/**
 * Get the index of the legal hold account in the To, Cc, or Bcc field.
 * @param {Office.EmailAddressDetails[]} recipients The recipients in the To, Cc, or Bcc field.
 * @returns {number} The index of the legal hold account.
 */
function getLegalHoldAccountIndex(recipients) {
  return recipients.findIndex((recipient) => (recipient.emailAddress).toLowerCase() === LEGAL_HOLD_ACCOUNT);
}

/**
 * Remove the legal hold email account from the To, Cc, or Bcc field of a message.
 * @param {Office.AddinCommands.Event} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
 * @param {Office.Recipients} recipientField The recipient object of the To, Cc, or Bcc field of a message.
 */
function removeLegalHoldAccount(event, recipientField) {
  recipientField.getAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve recipients from the field.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    const recipients = result.value;
    const legalHoldAccountIndex = getLegalHoldAccountIndex(recipients);
    if (legalHoldAccountIndex > -1) {
      recipients.splice(legalHoldAccountIndex, 1);
      recipientField.setAsync(recipients, { asyncContext: event }, (result) => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to set the recipients.");
          console.log(`Error: ${result.error.message}`);
          event.completed();
          return;
        }
    
        console.log(`${LEGAL_HOLD_ACCOUNT} has been removed.`);
        event.completed();
      });
    }
  });
}

/**
 * Add the legal hold email account to the Bcc field.
 * @param {Office.AddinCommands.Event} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
 * @param {Office.Recipients} recipientField The recipient object of the Bcc field.
 */
function addLegalHoldAccount(event, recipientField) {
  recipientField.getAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve recipients from the field.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    const recipients = result.value;
    const legalHoldAccountIndex = getLegalHoldAccountIndex(recipients);
    if (legalHoldAccountIndex === -1) {
      recipientField.addAsync([LEGAL_HOLD_ACCOUNT], { asyncContext: event }, (result) => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log(`Unable to add ${LEGAL_HOLD_ACCOUNT} as a recipient.`);
          console.log(`Error: ${result.error.message}`);
          event.completed();
          return;
        }

        console.log(`${LEGAL_HOLD_ACCOUNT} has been added to the Bcc field.`);
        event.completed();
      });
    }
  });
}

/**
 * Get the unique identifier (GUID) of a sensitivity label.
 * @param {string} sensitivityLabel The name of a sensitivity label.
 * @param {Office.SensitivityLabelDetails[]} sensitivityLabelCatalog The catalog of sensitivity labels.
 * @returns {number} The GUID of a sensitivity label. 
 */
function getLabelId(sensitivityLabel, sensitivityLabelCatalog) {
  return (sensitivityLabelCatalog.find((label) => label.name === sensitivityLabel)).id;
}

/**
 * Check if a member of the legal team is a recipient in the To, Cc, or Bcc field.
 * @param {Office.EmailAddressDetails[]} recipients The recipients in the To, Cc, or Bcc field.
 * @returns {boolean} Returns true if a member of the legal team is a recipient.
 */
function containsLegalTeamMember(recipients) {
  for (let i = 0; i < recipients.length; i++) {
    const emailAddress = recipients[i].emailAddress.toLowerCase();
    if (emailAddress.includes(LEGAL_TEAM_ACCOUNT_SUFFIX)) {
      return true;
    }
  }

  return false;
}

/**
 * Maps the event handler name specified in the manifest to its JavaScript counterpart.
 */
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onSensitivityLabelChangedHandler", onSensitivityLabelChangedHandler);
