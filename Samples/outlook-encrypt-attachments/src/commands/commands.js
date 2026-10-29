/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
*/

Office.onReady();

const demoXorKey = "sample only - not secure";
const encryptedAttachmentPrefix = "encrypted_";
const decryptedAttachmentPrefix = "decrypted_";
var totalOptionalAttendees = 0;
var totalRequiredAttendees = 0;
var totalDistributionLists = 0;

/**
 * Runs when the user creates a new appointment as the organizer.
 * @param {Office.MailboxEvent} event The event-based activation event.
 */
function onAppointmentComposeHandler(event) {
  // Call event.completed() after all asynchronous operations finish.

  console.log("onAppointmentComposeHandler(): entered!");

  let originalAppointmentDate = {}; // Store the original dates and times in sessionData.

  Office.context.mailbox.item.start.getAsync((asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`onAppointmentComposeHandler(): Action failed with message ${asyncResult.error.message}`);
      event.completed();
      return;
    }
    console.log(`onAppointmentComposeHandler(): Appointment starts: ${asyncResult.value}`);
    originalAppointmentDate.start = asyncResult.value;

    Office.context.mailbox.item.end.getAsync((asyncResult2) => {
      if (asyncResult2.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`onAppointmentComposeHandler(): Action failed with message ${asyncResult2.error.message}`);
        event.completed();
        return;
      }
      originalAppointmentDate.end = asyncResult2.value;

      console.log(`onAppointmentComposeHandler(): Appointment ends: ${asyncResult2.value}`);
      console.log("onAppointmentComposeHandler(): Setting SessionData...");

      Office.context.mailbox.item.sessionData.setAsync("appointment_info", JSON.stringify(originalAppointmentDate), function(asyncResult3){
        if (asyncResult3.status === Office.AsyncResultStatus.Succeeded) {
          console.log("onAppointmentComposeHandler(): sessionData set");
          // Add a notification that prompts the user to open the task pane for sample instructions.

          console.log("onAppointmentComposeHandler(): Adding notification message...");

          // The actions array applies only to insightMessage notifications.
          Office.context.mailbox.item.notificationMessages.addAsync(
            "showInfoBarForSampleInstructions",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
              message:
                "Open the task pane for details about running the Outlook Event-based Activation Sample Add-in.",
              icon: "Icon.16x16",
              actions: [
                {
                  actionText: "Show task pane",
                  actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
                  commandId: "appOrgTaskPaneButton",
                  contextData: "{''}",
                },
              ],
            },
            function () {
              console.log(
                "onAppointmentComposeHandler(): Office.context.mailbox.item.notificationMessages.addAsync completed"
              );
              event.completed();
            }
          );
        }
        else {
          console.error(`onAppointmentComposeHandler(): Action failed with message ${asyncResult3.error.message}`);
          event.completed();
        }
      });
    });
  });
}
/**
 * Runs when the user composes a new message.
 * @param {Office.MailboxEvent} event The event-based activation event.
 */
function onMessageComposeHandler(event) {
  console.log("onMessageComposeHandler(): entered!");

  // Add a notification that prompts the user to open the task pane for sample instructions.
  // The actions array applies only to insightMessage notifications.
  var options = { asyncContext: { callingEvent: event } };
  Office.context.mailbox.item.notificationMessages.addAsync(
    "showInfoBarForSampleInstructions",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
      message: "Open the task pane for details about running the Outlook Event-based Activation Sample Add-in.",
      icon: "Icon.16x16",
      actions: [
        {
          actionText: "Show task pane",
          actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
          commandId: "msgComposeOpenPaneButton",
          contextData: "{''}",
        },
      ],
    },
    options,
    function (result) {
      console.log("onMessageComposeHandler(): Notification message added");
      result.asyncContext.callingEvent.completed();
    }
  );
}
/**
 * Runs when an attendee is added to or removed from an appointment.
 * @param {Office.MailboxEvent} event The event-based activation event.
 */
function onAppointmentAttendeesChangedHandler(event) {

  // Call event.completed() after all asynchronous operations finish.

  var totalOptionalAttendees = 0;
  var totalRequiredAttendees = 0;
  var totalDistributionLists = 0;

  console.log(`onAppointmentAttendeesChangedHandler(): type = ${event.type}; requiredAttendees: ${event.changedRecipientFields.requiredAttendees}; optionalAttendees: ${event.changedRecipientFields.optionalAttendees}; resources: ${event.changedRecipientFields.resources};`);

  // Use sequential asynchronous Office API calls to count required attendees, optional attendees, and distribution lists.
  // Add, update, or remove notifications based on the counts.

  // Get the required attendees.
  Office.context.mailbox.item.requiredAttendees.getAsync(function (asyncResult) {
    console.log(`onAppointmentAttendeesChangedHandler(): getAsync => asyncResult.status: ${asyncResult.status}`);
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      // Count the required attendees without including the current user.
      var apptRequiredAttendees = asyncResult.value;

      for (var i = 0; i < apptRequiredAttendees.length; i++) {
        console.log(
          `onAppointmentAttendeesChangedHandler(): Required attendee: ${apptRequiredAttendees[i].displayName} (${apptRequiredAttendees[i].emailAddress}) - type: ${apptRequiredAttendees[i].recipientType}`
        );
      }

      // Exclude the current user, who can appear in the attendee array in Outlook on Windows.
      console.log(`onAppointmentAttendeesChangedHandler(): Filtering out attendees for current user ${Office.context.mailbox.userProfile.emailAddress}...`);
      var apptRequiredAttendeesWithoutUser = apptRequiredAttendees.filter(function(attendee){return attendee.emailAddress !== Office.context.mailbox.userProfile.emailAddress;});

      totalRequiredAttendees = apptRequiredAttendeesWithoutUser.length;
      currentDistributionLists = apptRequiredAttendees.filter(function(attendee){return attendee.recipientType === "distributionList";});
      if (currentDistributionLists.length !== 0){
        totalDistributionLists = totalDistributionLists + currentDistributionLists.length;
      }

      console.log(`onAppointmentAttendeesChangedHandler(): totalRequiredAttendees = ${totalRequiredAttendees}; totalDistributionLists = ${totalDistributionLists}`);

      // Get the optional attendees.
      Office.context.mailbox.item.optionalAttendees.getAsync(function (asyncResult2) {
        console.log(`onAppointmentAttendeesChangedHandler(): getAsync => asyncResult2.status: ${asyncResult2.status}`);
        if (asyncResult2.status === Office.AsyncResultStatus.Succeeded) {
          // Count the optional attendees without including the current user.
          var apptOptionalAttendees = asyncResult2.value;
          // Exclude the current user, who can appear in the attendee array in Outlook on Windows.
          var apptOptionalAttendeesWithoutUser = apptOptionalAttendees.filter(function(attendee){return attendee.emailAddress !== Office.context.mailbox.userProfile.emailAddress;});

          totalOptionalAttendees = apptOptionalAttendeesWithoutUser.length;
          currentDistributionLists = apptOptionalAttendees.filter(function(attendee){return attendee.recipientType === "distributionList";});
          if (currentDistributionLists.length !== 0){
            totalDistributionLists = totalDistributionLists + currentDistributionLists.length;
          }

        } else {
          console.error(`onAppointmentAttendeesChangedHandler(): Error with item.optionalAttendees.getAsync(): ${asyncResult2.error}`);
        }

        console.log(`onAppointmentAttendeesChangedHandler(): totalRequiredAttendees = ${totalRequiredAttendees}; totalOptionalAttendees = ${totalOptionalAttendees}; totalDistributionLists = ${totalDistributionLists}`);

        // Update the notifications based on the attendee counts.
        if (totalOptionalAttendees === 0 && totalRequiredAttendees === 0) {
          // Remove the attendee count notification when no attendees remain.
          Office.context.mailbox.item.notificationMessages.removeAsync(
            "attendeesChanged",
            null,
            function (asyncResult3) {
              console.log(`onAppointmentAttendeesChangedHandler() removeAsync:attendeesChanged => asyncResult3.status: ${asyncResult3.status}`);
              // Remove the distribution list warning if distribution lists are still counted.
              if (asyncResult3.status === Office.AsyncResultStatus.Succeeded && totalDistributionLists !== 0) {
                Office.context.mailbox.item.notificationMessages.removeAsync(
                  "distributionListWarning",
                  null,
                  function (asyncResultDLs) {
                    console.log(`onAppointmentAttendeesChangedHandler() removeAsync:distributionListWarning => asyncResultDLs.status: ${asyncResultDLs.status}`);
                    event.completed(); // Complete the event after the notification update.
                  }
                );
              } else {
                // Complete the event when no further notification updates are needed.
                event.completed();
              }
            }
          );
        } else
        {
          // Update the attendee count notification.
          Office.context.mailbox.item.notificationMessages.replaceAsync(
            "attendeesChanged",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
              message: `Your appointment has ${totalRequiredAttendees} required and ${totalOptionalAttendees} optional attendees`,
              icon: "Icon.16x16",
              persistent: false,
            },
            function (asyncResult4) {
              console.log(`onAppointmentAttendeesChangedHandler(): replaceAsync:attendeesChanged => asyncResult4.status: ${asyncResult4.status}`);
              if (asyncResult4.status === Office.AsyncResultStatus.Succeeded) {
                var anyExistingWarningMessages = false;

                Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult5) {
                  console.log(`onAppointmentAttendeesChangedHandler(): getAllAsync(): asyncResult5.status = ${asyncResult5.status}`);
                  if (asyncResult5.status === Office.AsyncResultStatus.Succeeded) {
                    let distributionListWarningMessages = asyncResult5.value.filter(message => message.key === "distributionListWarning");
                    if (distributionListWarningMessages.length !== 0)
                    {
                      anyExistingWarningMessages = true;
                    }
                  }
                  else {
                    console.error(`onAppointmentAttendeesChangedHandler(): Error with item.notificationMessages.getAllAsync(): ${asyncResult5.error}`);
                  }

                  if (totalDistributionLists === 0 && anyExistingWarningMessages === true) {
                    // Remove the warning when no distribution lists remain.
                    Office.context.mailbox.item.notificationMessages.removeAsync(
                      "distributionListWarning",
                      null,
                      function (asyncResult6) {
                        console.log(`onAppointmentAttendeesChangedHandler() removeAsync:distributionListWarning => asyncResult6.status: ${asyncResult6.status}`);
                        if (asyncResult6.status === Office.AsyncResultStatus.Succeeded) {;
                          event.completed();
                          return;
                        } else {
                          console.error(`onAppointmentAttendeesChangedHandler(): Error with item.notificationMessages.removeAsync(): ${asyncResult6.error}`);
                        }
                      }
                    );
                  }

                  if (totalDistributionLists !== 0) {
                    var warningMessage;
                    if (totalDistributionLists === 1) {
                      warningMessage = `Warning! Your appointment has a distribution list! Make sure you have chosen the correct one!`;
                    } else {
                      warningMessage = `Warning! Your appointment has ${totalDistributionLists} distribution lists! Make sure you have chosen the correct one!`;
                    }

                    // Update the distribution list warning.
                    Office.context.mailbox.item.notificationMessages.replaceAsync(
                      "distributionListWarning",
                      {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: warningMessage,
                        icon: "Icon.16x16",
                        persistent: false
                      },
                      function (asyncResult7) {
                        console.log(`onAppointmentAttendeesChangedHandler(): replaceAsync:distributionListWarning => asyncResult7.status: ${asyncResult7.status}`);
                        event.completed(); // Complete the event after the notification update.
                      }
                    );
                  }
                  else{
                    event.completed(); // Complete the event when no warning is needed.
                  }
                });
              }
              else {
                console.error(`onAppointmentAttendeesChangedHandler(): Error with item.notificationMessages.replaceAsync(): ${asyncResult4.error}`);
              }
            }
          );
        }
      });
    } else {
      console.error(`onAppointmentAttendeesChangedHandler(): Unexpected: asyncResult.status = ${asyncResult.status}`);
      event.completed(); // Complete the event if the attendee lookup fails.
    }
  });
}
/**
 * Runs when an attachment is added to or removed from a message or appointment being composed.
 * @param {Office.MailboxEvent} event The event-based activation event.
 */
function onItemAttachmentsChangedHandler(event) {
  if (event.attachmentStatus === "removed"){
    console.log("onItemAttachmentsChangedHandler(): Not processing removed attachments; leaving...");
    event.completed(); // Complete the event without processing removed attachments.
    return;
  }

  console.log(`onItemAttachmentsChangedHandler(): ${event.attachmentDetails.name} (${event.attachmentStatus})`);

  const fileName = event.attachmentDetails.name;
  if (fileName.startsWith(encryptedAttachmentPrefix) || fileName.startsWith(decryptedAttachmentPrefix)) {
    event.completed();
    return;
  }

  const item = Office.context.mailbox.item;
  const context = { fileName: fileName, callingEvent: event };
  // Store the processing marker in sessionData to enforce the one-attachment limit across event runtime activations.
  item.sessionData.getAllAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      failAttachmentProcessing(context, result.error.message);
      return;
    }
    if (result.value.attachmentDemoStarted === "true") {
      console.log("Attachment demonstration already started for this compose session.");
      event.completed();
      return;
    }
    item.sessionData.setAsync("attachmentDemoStarted", "true", function (setResult) {
      if (setResult.status !== Office.AsyncResultStatus.Succeeded) {
        failAttachmentProcessing(context, setResult.error.message);
        return;
      }
      context.sessionMarked = true;
      item.getAttachmentContentAsync(event.attachmentDetails.id, { asyncContext: context }, handleAttachmentsCallback);
    });
  });
}
/**
 * Applies a reversible, insecure XOR transformation to Base64-encoded bytes.
 * Applying the transformation twice restores the input.
 * This is a demonstration only, not encryption suitable for production.
 */
function transformDemoAttachment(base64) {
  if (typeof base64 !== "string" ||
      !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(base64)) {
    throw new Error("Attachment content is not valid Base64.");
  }
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const output = [];
  let byteIndex = 0;
  for (let i = 0; i < base64.length; i += 4) {
    const first = alphabet.indexOf(base64[i]);
    const second = alphabet.indexOf(base64[i + 1]);
    const third = base64[i + 2] === "=" ? 0 : alphabet.indexOf(base64[i + 2]);
    const fourth = base64[i + 3] === "=" ? 0 : alphabet.indexOf(base64[i + 3]);
    const bytes = [(first << 2) | (second >> 4)];
    if (base64[i + 2] !== "=") {
      bytes.push(((second & 15) << 4) | (third >> 2));
    }
    if (base64[i + 3] !== "=") {
      bytes.push(((third & 3) << 6) | fourth);
    }
    for (let j = 0; j < bytes.length; j++) {
      bytes[j] ^= demoXorKey.charCodeAt(byteIndex++ % demoXorKey.length);
    }
    output.push(
      alphabet[bytes[0] >> 2],
      alphabet[((bytes[0] & 3) << 4) | (bytes.length > 1 ? bytes[1] >> 4 : 0)],
      bytes.length > 1 ? alphabet[((bytes[1] & 15) << 2) | (bytes.length > 2 ? bytes[2] >> 6 : 0)] : "=",
      bytes.length > 2 ? alphabet[bytes[2] & 63] : "="
    );
  }
  return output.join("");
}

function failAttachmentProcessing(context, message, removeProgress = false) {
  console.error(`Attachment demonstration failed: ${message}`);
  const item = Office.context.mailbox.item;
  function showError() {
    item.notificationMessages.replaceAsync("attachmentDemoError", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: `Attachment demo failed: ${message}`.slice(0, 150),
    }, function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Displaying attachment error failed: ${result.error.message}`);
      }
      context.callingEvent.completed();
    });
  }
  function releaseSession() {
    if (context.sessionMarked) {
      item.sessionData.removeAsync("attachmentDemoStarted", function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(`Resetting attachment demonstration failed: ${result.error.message}`);
        }
        showError();
      });
    } else {
      showError();
    }
  }
  if (removeProgress) {
    item.notificationMessages.removeAsync("processingAttachments", function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Removing attachment progress notification failed: ${result.error.message}`);
      }
      releaseSession();
    });
  } else {
    releaseSession();
  }
}
/**
 * Demonstrates a file data transformation. Cloud, email message, and calendar item attachments are not processed.
 * @param {Office.AsyncResult} result The attachment content and processing context.
 */
function handleAttachmentsCallback(result) {
  if (result.status !== Office.AsyncResultStatus.Succeeded) {
    failAttachmentProcessing(result.asyncContext, result.error.message);
    return;
  }
  const context = result.asyncContext;
  const fileName = context.fileName;
  console.log(`handleAttachmentsCallback(): result.value.format = ${result.value.format}`);
  // To log the Base64 content, uncomment the following line. Do not use console.dir() with runtime logging enabled.
  // console.dir(result.value.content);

  // Process the attachment according to its content format.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // The progress notification is removed after the decrypted attachment is added and may appear only briefly.
      var options = { asyncContext: context };
      const base64 = result.value.content;
      Office.context.mailbox.item.notificationMessages.addAsync("processingAttachments", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
        message: `Please wait while the '${fileName}' attachment is encrypted...`,
      }, options, function(asyncResult){
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            let ciphertext;
            try {
              ciphertext = transformDemoAttachment(base64);
            } catch (ex) {
              failAttachmentProcessing(context, ex.message, true);
              return;
            }
            encryptAttachment(ciphertext, context);
          } else {
            failAttachmentProcessing(context, asyncResult.error.message);
          }
      });
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Skip email message attachments.
      console.log("handleAttachmentsCallback(): Attachment is a message.");
      result.asyncContext.callingEvent.completed();
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Skip calendar item attachments.
      console.log("handleAttachmentsCallback(): Attachment is a calendar item.");
      result.asyncContext.callingEvent.completed();
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Skip cloud attachments.
      console.log("handleAttachmentsCallback(): Attachment is a cloud attachment.");
      result.asyncContext.callingEvent.completed();
      break;
    default:
      // Skip unsupported attachment formats.
      console.warn("handleAttachmentsCallback(): Not handling unsupported attachment.");
      result.asyncContext.callingEvent.completed();
      break;
  }
}
/**
 * Adds the Base64-encoded XOR demonstration output as an attachment.
 * @param {string} encryptedData The Base64-encoded transformation output.
 * @param {object} context The attachment name, calling event, and session marker state.
 */
function encryptAttachment(encryptedData, context) {
  const fileName = context.fileName;
  console.log(`encryptAttachment(): Encrypting file '${fileName}'...`);
  // To log the transformed data, uncomment the following line.
  // console.dir(encryptedData);

  var base64EncryptedData = encryptedData;
  var encryptedFileName = `${encryptedAttachmentPrefix}${fileName}`;
  var options = { asyncContext: context, isInline: false };

  // To log the Base64-encoded transformed data, uncomment the following lines.
  // console.log("encryptAttachment(): base64 encrypted data:");
  // console.dir(base64EncryptedData);

  console.log(`encryptAttachment(): Adding encrypted file '${encryptedFileName}'...`);
  Office.context.mailbox.item.addFileAttachmentFromBase64Async(
    base64EncryptedData,
    encryptedFileName,
    options,
    function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        failAttachmentProcessing(context, asyncResult.error.message, true);
        return;
      }
      options = { asyncContext: { fileName: fileName, encryptedFileName: encryptedFileName, callingEvent: context.callingEvent, encryptedData: encryptedData, sessionMarked: context.sessionMarked } };
      console.log(`encryptAttachment(): Added encrypted attachment '${encryptedFileName}'; now decrypting...`);
      // To log the attachment operation result, uncomment the following line.
      // console.dir(asyncResult);
      decryptAttachment(options);
    }
  );
}
/**
 * Reverses the XOR demonstration and attaches the restored file.
 * @param {object} result An options object whose asyncContext property contains the attachment data and calling event.
 */
function decryptAttachment(result) {
  const fileName = result.asyncContext.fileName;
  console.log(`decryptAttachment(): Decrypting file '${result.asyncContext.encryptedFileName}'...`);
  let originalText;
  try {
    originalText = transformDemoAttachment(result.asyncContext.encryptedData);
  } catch (ex) {
    failAttachmentProcessing(result.asyncContext, ex.message, true);
    return;
  }
  var decryptedFileName = `${decryptedAttachmentPrefix}${fileName}`;

  // To log the restored Base64 content, uncomment the following line.
  // console.log(`decryptAttachment(): Original base64: ${originalText}`);
  console.log(`decryptAttachment(): Adding decrypted file '${decryptedFileName}'...`);

  var options = { asyncContext: result.asyncContext };
  Office.context.mailbox.item.addFileAttachmentFromBase64Async(
    originalText,
    decryptedFileName,
    options,
    function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        failAttachmentProcessing(result.asyncContext, asyncResult.error.message, true);
        return;
      }
      console.log(`decryptAttachment(): Added decrypted attachment '${decryptedFileName}'`);
      // To log the attachment operation result, uncomment the following line.
      // console.dir(asyncResult);
      Office.context.mailbox.item.notificationMessages.removeAsync("processingAttachments", options, function (asyncResult2) {
        if (asyncResult2.status !== Office.AsyncResultStatus.Succeeded) {
          failAttachmentProcessing(result.asyncContext, asyncResult2.error.message);
          return;
        }
        console.log("decryptAttachment(): Notification message removed.");
        Office.context.mailbox.item.notificationMessages.addAsync("attachmentsAdded", {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: `The '${fileName}' attachment has been encrypted and decrypted and added as reference attachments for your review.`,
          icon: "Icon.16x16",
          persistent: false,
        }, options, function(asyncResult3){
          if (asyncResult3.status !== Office.AsyncResultStatus.Succeeded) {
            failAttachmentProcessing(result.asyncContext, asyncResult3.error.message);
            return;
          }
          console.log("decryptAttachment(): All attachment processing operations completed!");
          asyncResult3.asyncContext.callingEvent.completed();
        });
      });
    }
  );
}

/**
 * Runs when the user changes the date or time of an appointment.
 * @param {Office.MailboxEvent} event The event-based activation event.
 */
function onAppointmentTimeChangedHandler(event) {
  console.log(`onAppointmentTimeChangedHandler(): type: ${event.type}; start: ${event.start}; end: ${event.end}`);

  Office.context.mailbox.item.sessionData.getAsync("appointment_info", function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(`onAppointmentTimeChangedHandler(): sessionData retrieved: ${asyncResult.value}`);

      try {
        // Convert the original start and end dates and times to the client's local time.
        let originalAppointmentDate = JSON.parse(asyncResult.value);
        let originalAppointmentDateStartDate = new Date(originalAppointmentDate.start);
        let originalAppointmentDateStartEnd = new Date(originalAppointmentDate.end);
        let convertedToLocalStartUtc = Office.context.mailbox.convertToLocalClientTime(
          originalAppointmentDateStartDate
        );
        let convertedToLocalEndUtc = Office.context.mailbox.convertToLocalClientTime(originalAppointmentDateStartEnd);
        let convertedToLocalStart = new Date(
          convertedToLocalStartUtc.year,
          convertedToLocalStartUtc.month,
          convertedToLocalStartUtc.date,
          convertedToLocalStartUtc.hours,
          convertedToLocalStartUtc.minutes
        );
        let convertedToLocalEnd = new Date(
          convertedToLocalEndUtc.year,
          convertedToLocalEndUtc.month,
          convertedToLocalEndUtc.date,
          convertedToLocalEndUtc.hours,
          convertedToLocalEndUtc.minutes
        );

        // Add a notification with the original start and end dates and times if the notification does not already exist.
        Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult2) {
          console.log(`getAllAsync(): asyncResult.status = ${asyncResult2.status}`);
          if (asyncResult2.status != "failed") {
            for (let index = 0; index < asyncResult2.value.length; index++) {
              const element = asyncResult2.value[index];
              if (element.key === "timeChanged") {
                // Preserve the existing notification so it continues to show the original dates and times.
                event.completed();
                return;
              }
            }
          }

          var originalDateMessage = `Original date/time: Start = ${convertedToLocalStart.toLocaleDateString()} ${convertedToLocalStart.toLocaleTimeString()}; End = ${convertedToLocalEnd.toLocaleDateString()} ${convertedToLocalEnd.toLocaleTimeString()}`;
          Office.context.mailbox.item.notificationMessages.replaceAsync(
            "timeChanged",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
              message: originalDateMessage,
              icon: "Icon.16x16",
              persistent: false,
            },
            function (asyncResult3) {
              console.log("replaceAsync() for 'timeChanged' completed");
              event.completed();
            }
          );
        });
      }
      catch (ex) {
        console.error(`onAppointmentTimeChangedHandler(): Error! ${ex}`);
        event.completed();
      }
    } else {
      console.error(`onAppointmentTimeChangedHandler(): Action failed with message ${asyncResult.error.message}`);
      event.completed();
    }
  });
}

// Map each manifest action ID (unified manifest) or LaunchEvent FunctionName (add-in only manifest) to its handler.
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
Office.actions.associate("onMessageAttachmentsChangedHandler", onItemAttachmentsChangedHandler);
Office.actions.associate("onAppointmentAttachmentsChangedHandler", onItemAttachmentsChangedHandler);
Office.actions.associate("onAppointmentAttendeesChangedHandler", onAppointmentAttendeesChangedHandler);
Office.actions.associate("onAppointmentTimeChangedHandler", onAppointmentTimeChangedHandler);
