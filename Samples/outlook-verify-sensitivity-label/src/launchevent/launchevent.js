/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * In classic Outlook on Windows, when the event handler runs, code in Office.onReady() or Office.initialize isn't run.
 * Add any startup logic needed by handlers to the event handler itself.
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
 * Handle the OnMessageRecipientsChanged event by checking whether the legal hold email account or a -legal@fabrikam.com
 * address was added to the To, Cc, or Bcc field of a message. If the legal hold account was added to the To or Cc field,
 * the account is removed. If the legal hold account was added to the Bcc field, it's only removed if the sensitivity label
 * of a message isn't set to Highly Confidential. When a -legal@fabrikam.com address is present in any of the recipient fields,
 * the message is checked for the Highly Confidential sensitivity label.
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged event object.
 */
function onMessageRecipientsChangedHandler(event) {
  if (event.changedRecipientFields.to) {
    Office.context.mailbox.item.to.getAsync({ asyncContext: event }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to get the recipients from the To field.");
        console.log(`Error: ${result.error.message}`);
        event.completed();
        return;
      }

      const recipients = result.value;
      if (containsLegalTeamMember(recipients)) {
        ensureHighlyConfidentialLabelSet(event, () => {
          removeLegalHoldAccount(event, recipients, Office.context.mailbox.item.to);
        });
      } else {
        removeLegalHoldAccount(event, recipients, Office.context.mailbox.item.to);
      }
    });
  }

  if (event.changedRecipientFields.cc) {
    Office.context.mailbox.item.cc.getAsync({ asyncContext: event }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to get the recipients from the Cc field.");
        console.log(`Error: ${result.error.message}`);
        event.completed();
        return;
      }

      const recipients = result.value;
      if (containsLegalTeamMember(recipients)) {
        ensureHighlyConfidentialLabelSet(event, () => {
          removeLegalHoldAccount(event, recipients, Office.context.mailbox.item.cc);
        });
      } else {
        removeLegalHoldAccount(event, recipients, Office.context.mailbox.item.cc);
      }
    });
  }

  if (event.changedRecipientFields.bcc) {
    Office.context.mailbox.item.bcc.getAsync({ asyncContext: event }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to get the recipients from the Bcc field.");
        console.log(`Error: ${result.error.message}`);
        event.completed();
        return;
      }

      const recipients = result.value;
      if (containsLegalTeamMember(recipients)) {
        ensureHighlyConfidentialLabelSet(event, () => {
          checkForLegalHoldAccount(event);
        });
      } else {
        checkForLegalHoldAccount(event);
      }
    });
  }
}

/**
 * Handle the OnMessageAttachmentsChanged event by checking whether the message contains an attachment.
 * If the message has at least one attachment, the event handler checks for the Highly Confidential sensitivity
 * label on the message and sets it if needed.
 * @param {Office.MailboxEvent} event The OnMessageAttachmentsChanged event object.
 */
function onMessageAttachmentsChangedHandler(event) {
  Office.context.mailbox.item.getAttachmentsAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to check for attachments.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    const attachments = result.value;
    if (attachments.length > 0) {
      ensureHighlyConfidentialLabelSet(event);
    } else {
      event.completed();
    }
  });
}

/**
 * Handle the OnMessageSend event by displaying a dialog when the legal hold account is added to the Bcc field and the
 * Highly Confidential sensitivity label is set on the message. If these conditions aren't met, the message is sent.
 * @param {Office.MailboxEvent} event The OnMessageSend event object.
 */
function onMessageSendHandler(event) {
  Office.context.mailbox.item.bcc.getAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to get the recipients from the Bcc field.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    const legalHoldAccountIndex = getLegalHoldAccountIndex(result.value);
    if (legalHoldAccountIndex === -1) {
      event.completed({ allowEvent: true });
    } else {
      getSensitivityLabel(event, (currentLabel, labelCatalog) => {
        const highlyConfidentialLabel = getLabel("Highly Confidential", labelCatalog);
        let labelId = highlyConfidentialLabel.id;

        if (highlyConfidentialLabel.children.length > 0) {
          labelId = highlyConfidentialLabel.children[0].id;
        }

        if (currentLabel != labelId) {
          event.completed({ allowEvent: true });
        } else {
          event.completed({ allowEvent: false, errorMessage: "Due to the contents of your message, the sensitivity label has been set to Highly Confidential and the Legal Hold account has been added to the Bcc field.\nTo learn more, see Fabrikam's information protection policy.\n\nDo you need to make changes to your message?" });
        }
      });
    }
  });
}

/**
 * Handle the OnSensitivityLabelChanged event by verifying that Highly Confidential label is added when there's
 * at least one attachment or a recipient who's a member of the legal team.
 * @param {Office.MailboxEvent} event The OnSensitivityLabelChanged event object.
 */
function onSensitivityLabelChangedHandler(event) {
  Office.context.mailbox.item.getAttachmentsAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to check for attachments.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    let attachmentPresent;
    if (result.value.length > 0) {
      attachmentPresent = true;
      evaluateConditions(event, undefined, attachmentPresent);
    } else {
      attachmentPresent = false;
      Office.context.mailbox.item.to.getAsync({ asyncContext: { event: event, attachmentPresent: attachmentPresent } }, (result) => {
        const event = result.asyncContext.event;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to get the recipients from the To field.");
          console.log(`Error: ${result.error.message}`);
          event.completed();
          return;
        }
    
        let hasLegalTeamMember = false;
        if (containsLegalTeamMember(result.value)) {
          hasLegalTeamMember = true;
          evaluateConditions(event, hasLegalTeamMember, result.asyncContext.attachmentPresent);
        } else {
          Office.context.mailbox.item.cc.getAsync({ asyncContext: { event: event, attachmentPresent: attachmentPresent } }, (result) => {
            const event = result.asyncContext.event;
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.log("Unable to get the recipients from the Cc field.");
              console.log(`Error: ${result.error.message}`);
              event.completed();
              return;
            }
    
            if (containsLegalTeamMember(result.value)) {
              hasLegalTeamMember = true;
              evaluateConditions(event, hasLegalTeamMember, result.asyncContext.attachmentPresent);
            } else {
              Office.context.mailbox.item.bcc.getAsync({ asyncContext: { event: event, attachmentPresent: attachmentPresent } }, (result) => {
                const event = result.asyncContext.event;
                if (result.status === Office.AsyncResultStatus.Failed) {
                  console.log("Unable to get the recipients from the Bcc field.");
                  console.log(`Error: ${result.error.message}`);
                  event.completed();
                  return;
                }
    
                if (containsLegalTeamMember(result.value)) {
                  hasLegalTeamMember = true;
                }

                evaluateConditions(event, hasLegalTeamMember, result.asyncContext.attachmentPresent);
              });
            }
          });
        }
      });
    }
  });
}

/**
 * Check whether the current message has an attachment or a recipient who's a member of the legal team.
 * If either of these conditions is true, ensure that the Highly Confidential sensitivity label is set and the
 * legal hold email account is in the Bcc field.
 * @param {Office.MailboxEvent} event The OnSensitivityLabelChanged event object.
 * @param {boolean} hasLegalTeamMember Indicates whether a member of the legal team is a recipient in the To, Cc, or Bcc field.
 *                                     Defaults to false if not specified.
 * @param {boolean} hasAttachment Indicates whether the message has an attachment. Defaults to false if not specified.
 */
function evaluateConditions(event, hasLegalTeamMember = false, hasAttachment = false) {
  if (hasLegalTeamMember === true || hasAttachment === true) {
    ensureHighlyConfidentialLabelSet(event, () => {
      checkForLegalHoldAccount(event);
    });
  } else {
    checkForLegalHoldAccount(event);
  }
}

/**
 * Get the current sensitivity label of a message.
 * @param {function} [callback] - Optional callback function to execute after getting the current sensitivity label.
 *                                If not provided, event.completed() will be called automatically.
 */
function getSensitivityLabel(event, callback) {
  Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve the status of the sensitivity label catalog.");
      console.log(`Error: ${result.error.message}`);
      if (callback) {
        callback();
      } else {
        event.completed();
      }
      return;
    }

    Office.context.sensitivityLabelsCatalog.getAsync({ asyncContext: event }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to retrieve the catalog of sensitivity labels.");
        console.log(`Error: ${result.error.message}`);
        if (callback) {
          callback();
        } else {
          event.completed();
        }
        return;
      }

      const sensitivityLabelCatalog = result.value;
      Office.context.mailbox.item.sensitivityLabel.getAsync({ asyncContext: { event: event, sensitivityLabelCatalog: sensitivityLabelCatalog } }, (result) => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to get the sensitivity label of the message.");
          console.log(`Error: ${result.error.message}`);
          if (callback) {
            callback();
          } else {
            event.completed();
          }
          return;
        }

        callback(result.value, result.asyncContext.sensitivityLabelCatalog);
      });
    });
  });
}

/**
 * Check that the Highly Confidential sensitivity label is set if a message contains an attachment or a recipient
 * who's a member of the legal team.
 * @param {Office.MailboxEvent} event The OnMessageAttachmentsChanged or OnSensitivityLabelChanged event object.
 * @param {function} [callback] - Optional callback function to execute after checking for the Highly Confidential label.
 *                                If not provided, event.completed() will be called automatically.
 */
function ensureHighlyConfidentialLabelSet(event, callback) {
  getSensitivityLabel(event, (currentLabel, labelCatalog) => {
    const highlyConfidentialLabel = getLabel("Highly Confidential", labelCatalog);
    let labelId = highlyConfidentialLabel.id;

    if (highlyConfidentialLabel.children.length > 0) {
      labelId = highlyConfidentialLabel.children[0].id;
    }

    if (currentLabel == labelId) {
      if (callback) {
        callback();
      } else {
        event.completed();
      }
    } else {
      Office.context.mailbox.item.sensitivityLabel.setAsync(labelId, { asyncContext: event }, (result) => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to set the Highly Confidential sensitivity label to the message.");
          console.log(`Error: ${result.error.message}`);
          if (callback) {
            callback();
          } else {
            event.completed();
          }
          return;
        }

        if (callback) {
          callback();
        } else {
          event.completed();
        }
      });
    }
  });
}

/**
 * Check whether the legal hold account was added to the Bcc field if the sensitivity label of a message is set to
 * Highly Confidential. If the account appears in the Bcc field, but the sensitivity label isn't set to
 * Highly Confidential, the account is removed from the message.
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
 */
function checkForLegalHoldAccount(event) {
  getSensitivityLabel(event, (currentLabel, labelCatalog) => {
    const highlyConfidentialLabel = getLabel("Highly Confidential", labelCatalog);
    let labelId = highlyConfidentialLabel.id;

    if (highlyConfidentialLabel.children.length > 0) {
      labelId = highlyConfidentialLabel.children[0].id;
    }

    if (currentLabel == labelId) {
      addLegalHoldAccount(event, Office.context.mailbox.item.bcc);
    } else {
      Office.context.mailbox.item.bcc.getAsync({ asyncContext: event }, (result) => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to get the recipients from the Bcc field.");
          console.log(`Error: ${result.error.message}`);
          event.completed();
          return;
        }
    
        const recipients = result.value;
        removeLegalHoldAccount(event, recipients, Office.context.mailbox.item.bcc);
      });
    }
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
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
 * @param {Office.EmailAddressDetails} recipients The array of recipients from the To, Cc, or Bcc field of a message.
 * @param {Office.Recipients} recipientField The recipient object of the To, Cc, or Bcc field of a message.
 */
function removeLegalHoldAccount(event, recipients, recipientField) {
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
}

/**
 * Add the legal hold email account to the Bcc field.
 * @param {Office.MailboxEvent} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
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
 * Get a sensitivity label from the catalog.
 * @param {string} sensitivityLabel The name of a sensitivity label.
 * @param {Office.SensitivityLabelDetails[]} sensitivityLabelCatalog The catalog of sensitivity labels.
 * @returns {number} The sensitivity label requested. 
 */
function getLabel(sensitivityLabel, sensitivityLabelCatalog) {
  return (sensitivityLabelCatalog.find((label) => label.name == sensitivityLabel));
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
Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
