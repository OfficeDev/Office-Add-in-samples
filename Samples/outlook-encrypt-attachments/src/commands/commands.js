/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
*/

//TEST In Outlook desktop
/* eslint-disable no-undef */ //For Office objects
// var CryptoJS = require("crypto-js"); //BUG Can't use: "Uncaught ReferenceError: require is not defined"
var fileName;
const secretKey = "secret key 123";
const encryptedAttachmentPrefix = "encrypted_";
const decryptedAttachmentPrefix = "decrypted_";
// eslint-disable-next-line no-unused-vars
Office.initialize = function (reason) {};
/**
 * Method that fires when a appointment is being created or edited
 * @param {Office.AsyncResult} result default: Office.AsyncResult
 */
function onAppointmentComposeHandler(event) {
  //NOTE: Must call event.completed() when any logic below is finished!

  console.log("onAppointmentComposeHandler(): entered!");

  let originalAppointmentDate = {}; //Create an object to cache the original date/time and persist it to localStorage

  Office.context.mailbox.item.start.getAsync((asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${asyncResult.error.message}`);
      event.completed();
      return;
    }
    console.log(`Appointment starts: ${asyncResult.value}`);
    originalAppointmentDate.start = asyncResult.value;

    Office.context.mailbox.item.end.getAsync((asyncResult2) => {
      if (asyncResult2.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${asyncResult2.error.message}`);
        event.completed();
        return;
      }

      console.log(`Appointment ends: ${asyncResult2.value}`);
      originalAppointmentDate.end = asyncResult2.value;
      localStorage.setItem("appointment_info", JSON.stringify(originalAppointmentDate));
      
      //NOTE: Clicking the "Show Task Pane" link in the InfoBar doesn't work. It is currently 'in backlog' status: https://github.com/OfficeDev/office-js/issues/2125
      //NOTE: actions array only applicable to insightMessage types
      //https://docs.microsoft.com/en-us/javascript/api/outlook/office.notificationmessages?view=outlook-js-preview
      Office.context.mailbox.item.notificationMessages.addAsync(
        "showInfoBarForSampleInstructions",
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage, //"insightMessage"
          message: "Open the Task Pane for details about running the Outlook Event-based Activation Sample Add-in",
          icon: "Icon.16x16",
          actions: [
            {
              actionText: "Show Task Pane",
              actionType: Office.MailboxEnums.ActionType.ShowTaskPane, //"showTaskPane"
              commandId: "appOrgTaskPaneButton",
              contextData: "{''}",
            },
          ],
        },
        function () {
          event.completed();
        }
      );
    });
  });
}
/**
 * Method that fires when an email is being created 
 * @param {Office.AsyncResult} result default: Office.AsyncResult
 */
function onMessageComposeHandler() {
  //Reminder to open the TaskPane
  showInfoBarForSampleInstructions();
}
/**
 * Method that fires when an attendee is added or removed from an appointment
 * @param {Office.AsyncResult} result default: Office.AsyncResult
 */
function onAppointmentAttendeesChangedHandler(event) {
  //https://docs.microsoft.com/en-us/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-preview&preserve-view=true

  //NOTE: Must call event.completed() when any logic below is finished!

  var totalOptionalAttendees = 0;
  var totalRequiredAttendees = 0;
  var totalDistributionLists = 0;

  console.log(`onAppointmentAttendeesChangedHandler() type = ${event.type}; changedRecipientFields = (dir dump on next line)`);
  console.dir(event.changedRecipientFields);

  Office.context.mailbox.item.requiredAttendees.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      var apptRequiredAttendees = asyncResult.value;
      //BUG 10/9/2021 On Outlook for Windows, when adding the first attendee it is also picking up a second attendee (possibly the sender)
      totalRequiredAttendees = apptRequiredAttendees.length;
      console.log(`totalRequiredAttendees = ${totalRequiredAttendees}`);
      currentDistributionLists = apptRequiredAttendees.filter(function(attendee){return attendee.recipientType === "distributionList";});
      if (currentDistributionLists.length !== 0){
        totalDistributionLists = totalDistributionLists + currentDistributionLists.length;
      }     

      Office.context.mailbox.item.optionalAttendees.getAsync(function (asyncResult2) {
        console.log(`status = ${asyncResult2.status}`);

        if (asyncResult2.status === Office.AsyncResultStatus.Succeeded) {
          var apptOptionalAttendees = asyncResult2.value;          
          totalOptionalAttendees = apptOptionalAttendees.length;
          currentDistributionLists = apptOptionalAttendees.filter(function(attendee){return attendee.recipientType === "distributionList";});
          if (currentDistributionLists.length !== 0){
            totalDistributionLists = totalDistributionLists + currentDistributionLists.length;
          }          
          
        } else {
          console.error(`Error with item.optionalAttendees.getAsync(): ${asyncResult2.error}`);
        }

        console.log(`totalDistributionLists = ${totalDistributionLists}`);

        if (totalOptionalAttendees === 0 && totalRequiredAttendees === 0) {
          //Remove the info bar with the recipients tally and any distribution list warnings if there are no longer any recipients
          Office.context.mailbox.item.notificationMessages.removeAsync(
            "attendeesChanged",
            null,
            function (asyncResult3) {
              if (asyncResult3.status === Office.AsyncResultStatus.Succeeded) {
                console.log(`asyncResult3.status = ${asyncResult3.status}`);

                Office.context.mailbox.item.notificationMessages.removeAsync(
                  "distributionListWarning",
                  null,
                  function (asyncResultDLs) {
                    if (asyncResultDLs.status === Office.AsyncResultStatus.Succeeded) {
                      console.log(`asyncResultDLs.status = ${asyncResultDLs.status}`);                      
                      event.completed(); //NOTE: Must call!                            
                    } else {
                      //REVIEW: This can happen if there are no more warning messages. Not sure if this is a logic bug or expected. Should consider making a notificationMessages.getAllAsync call like below?
                      console.error(`Error with item.notificationMessages.removeAsync(): ${asyncResultDLs.error}`);
                      event.completed(); //NOTE: Must call!                            
                    }
                  }
                );

              } else {
                console.error(`Error with item.notificationMessages.removeAsync(): ${asyncResult3.error}`);
                event.completed(); //NOTE: Must call!
              }
            }
          );
        } else 
        {
          Office.context.mailbox.item.notificationMessages.replaceAsync(
            "attendeesChanged",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
              message: `Your appointment has ${totalRequiredAttendees} required and ${totalOptionalAttendees} optional attendees`,
              icon: "Icon.16x16",
              persistent: false,
            },
            function (asyncResult4) {
              if (asyncResult4.status === Office.AsyncResultStatus.Succeeded) {  
                var anyExistingWarningMessages = false;
  
                Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult5) {
                  console.log(`getAllAsync(): asyncResult5.status = ${asyncResult5.status}`);
                  if (asyncResult5.status === Office.AsyncResultStatus.Succeeded) {
                    let distributionListWarningMessages = asyncResult5.value.filter(message => message.key === "distributionListWarning");
                    if (distributionListWarningMessages.length !== 0)
                    {
                      anyExistingWarningMessages = true;
                    }
                  }
                  else {
                    console.error(`Error with item.notificationMessages.getAllAsync(): ${asyncResult5.error}`);
                  }

                  if (totalDistributionLists === 0 && anyExistingWarningMessages === true) {                    
                    Office.context.mailbox.item.notificationMessages.removeAsync(
                      "distributionListWarning",
                      null,
                      function (asyncResult6) {
                        if (asyncResult6.status === Office.AsyncResultStatus.Succeeded) {;
                          event.completed();
                          return;
                        } else {
                          console.error(`Error with item.notificationMessages.removeAsync(): ${asyncResult6.error}`);
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

                    Office.context.mailbox.item.notificationMessages.replaceAsync(
                      "distributionListWarning",
                      {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: warningMessage,
                        icon: "Icon.16x16",
                        persistent: false
                      },
                      function () {
                        console.log("done");                      
                        event.completed(); //NOTE: Must call!
                      }
                    );
                  } 
                  else{                  
                    event.completed(); //NOTE: Must call! 
                  }
                });      
              }
              else {
                console.error(`Error with item.notificationMessages.replaceAsync(): ${asyncResult4.error}`);
              }
            }
          );
        }
      });
    } else {
      console.error(`Unexpected: asyncResult.status = ${asyncResult.status}`);
      asyncResult.completed(); //NOTE: Must call!
    }
  });
}
/**
 * Method that fires when an attachment is being added or removed from the compose email or appointment
 * @param {Office.AsyncResult} result default: Office.AsyncResult
 */
function onItemAttachmentsChangedHandler(event) {  
  //TODO Test adding two or more attachments at the same time!
  console.log("onItemAttachmentsChangedHandler: " + event.attachmentDetails.name + " (" + event.attachmentStatus + ")");
  if (event.attachmentDetails.name == `${decryptedAttachmentPrefix}${fileName}`) {
    //Don't process any more events - we've already encrypted the attachment and added it as another attachment, then decrypted that attachment and added it as well
    event.completed(); //NOTE: Must call!
    return;
  }
  if (fileName !== undefined){
    console.log("Skipping processing of further attachments - demo is done!");
    event.completed(); //NOTE: Must call!
    return;
  }

  //Process the first attachment. We'll encrypt it and add it as another attachment, then decrypted that attachment and add it as well
  fileName = event.attachmentDetails.name;
  var item = Office.context.mailbox.item;
  var options = { asyncContext: { currentItem: item } };
  item.getAttachmentsAsync(options, getAttachmentsCallback);
}
/**
 * Processes the attachments in the current item
 */
function getAttachmentsCallback(result) {
  if (result.value.length > 0) {
    for (i = 0; i < result.value.length; i++) {
      //https://docs.microsoft.com/en-us/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getAttachmentContentAsync_attachmentId__callback_
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}
/**
 * Method that encrypts base64 file data using CryptoJS and attaches the file to the email. Cloud, .eml and .ICalendar attachments will not be processed.
 * @param {string} result default: Office.AsyncResult
 */
function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  console.log(`handleAttachmentsCallback(): result.value.format = ${result.value.format}`);
  // console.dir(result.value.content); //NOTE: If you want to see the base64 data output to the console, uncomment this line

  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment

      //BUG Something is wrong with this flow in Outlook for Windows - nothing fires after adding the ProgressIndicator message
      //Set a notification message that we're processing the attachment. Note that this will be removed immediately after the decrypted attachment is added, and it may not be displayed for very long
      var options = { 'asyncContext': { base64: result.value.content } };
      Office.context.mailbox.item.notificationMessages.addAsync("processingAttachments", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
        message: `Please wait while the '${fileName}' attachment is encrypted...`,
      }, options, function(asyncResult){
          //Encrypt base64 file data using CryptoJS
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded){
            try {
               var ciphertext = CryptoJS.AES.encrypt(asyncResult.asyncContext.base64, secretKey).toString();              
              //TEST Using inline script; see https://stackoverflow.com/questions/62905663/how-to-import-crypto-js-in-either-a-vanilla-javascript-or-node-based-javascript. 10/11: Still the same issue
              //var ciphertext = encryptWithCrypto(asyncResult.asyncContext.base64, secretKey);

              //Then attaches the file to the email            
              console.log(`handleAttachmentsCallback(): starting processing of file '${fileName}'...`);
              encryptAttachment(ciphertext);                  
            }
            catch(ex){
              console.error(`handleAttachmentsCallback(): Error: ${ex}`);            
              Office.context.mailbox.item.notificationMessages.removeAsync("processingAttachments", function (result) {
        console.log("Notification message removed.");});
            }            
          }     
          else{
            console.error(`handleAttachmentsCallback(): Unexpected - status is ${asyncResult.status}`);            
          }     
      });      
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      console.log("Attachment is a message.");
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      console.log("Attachment is a calendar item.");
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      console.log("Attachment is a cloud attachment.");
      break;
    default:
      // Handle attachment formats that are not supported.
      console.log("Not handling unsupported attachment.");
      break;
  }
}
/**
 * Method that converts encrypted data to base64 and creates and adds a file attachment to the current email
 * @param {string} encryptedData default: "undefined"
 */
function encryptAttachment(encryptedData) {
  console.log(`encryptAttachment(): Encrypting file '${fileName}''...`);
  // console.dir(encryptedData); //NOTE: If you want to see the encrypted data output to the console, uncomment this line

  var base64EncryptedData = window.btoa(encryptedData);
  var encryptedFileName = `${encryptedAttachmentPrefix}${fileName}`;
  //NOTE: If you want to see the base64 data output to the console, uncomment these lines
  // console.log("encryptAttachment(): base64 encrypted data:");
  // console.dir(base64EncryptedData);

  console.log(`encryptAttachment(): Adding encrypted file '${encryptedFileName}'...`);
  Office.context.mailbox.item.addFileAttachmentFromBase64Async(
    base64EncryptedData,
    encryptedFileName,
    function (asyncResult) {
      console.log(`encryptAttachment(): Added encrypted attachment '${encryptedFileName}'; now decrypting...`);
      //console.dir(asyncResult); //NOTE: If you want to see the base64 data output to the console, uncomment this line
      decryptAttachment(encryptedData);
    }
  );
}
/**
 * Method that decrypts encrypted base64 file data using CryptoJS and attaches the file to the email
 * @param {string} encryptedData default: "undefined"
 */
function decryptAttachment(encryptedData) {
  console.log(`decryptAttachment(): Decrypting file '${fileName}''...`);

  var bytes = CryptoJS.AES.decrypt(encryptedData, secretKey);
  var originalText = bytes.toString(CryptoJS.enc.Utf8);
  var decryptedFileName = `${decryptedAttachmentPrefix}${fileName}`;

  // console.log(`decryptAttachment(): Original base64: ${originalText}`); //NOTE: If you want to see the base64 data output to the console, uncomment this line
  console.log(`decryptAttachment(): Adding decrypted file '${decryptedFileName}'...`);
  Office.context.mailbox.item.addFileAttachmentFromBase64Async(
    originalText,
    decryptedFileName,
    function (asyncResult) {
      console.log(`decryptAttachment(): Added decrypted attachment '${decryptedFileName}'`);
      // console.dir(asyncResult); //NOTE: If you want to see the base64 data output to the console, uncomment this line

      Office.context.mailbox.item.notificationMessages.removeAsync("processingAttachments", function (result) {
        console.log("Notification message removed.");        

        Office.context.mailbox.item.notificationMessages.addAsync("attachmentsAdded", {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: `The '${fileName}' attachment has been encrypted and decrypted and added as reference attachments for your review.`,
          icon: "Icon.16x16",
          persistent: false,
        });
      });
    }
  );
}
/**
 * Displays the InfoBar in the compose email. Purpose is to remind the user running this sample to open the Task Pane to display additional instructions and reference information
 */
function showInfoBarForSampleInstructions() {
  //NOTE: Clicking the "Show Task Pane" link in the InfoBar doesn't work. It is currently 'in backlog' status: https://github.com/OfficeDev/office-js/issues/2125
  //NOTE: actions array only applicable to insightMessage types
  //https://docs.microsoft.com/en-us/javascript/api/outlook/office.notificationmessages?view=outlook-js-preview
  Office.context.mailbox.item.notificationMessages.addAsync("showInfoBarForSampleInstructions", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage, //"insightMessage"
    message: "Open the Task Pane for details about running the Outlook Event-based Activation Sample Add-in",
    icon: "Icon.16x16",
    actions: [
      {
        actionText: "Show Task Pane",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane, //"showTaskPane"
        commandId: "msgComposeOpenPaneButton",
        contextData: "{''}",
      },
    ],
  });
}
/**
 * Method that fires when the user changes the date/time for an appointment
 * @param {string} result default: Office.AsyncResult
 */
function onAppointmentTimeChangedHandler(event) {
  //BUG Is this message preventing use of localStorage? Tracking Prevention blocked access to storage for https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
  console.dir(event);
  console.dir(event.type);
  console.dir(event.start);
  console.dir(event.end);

  let originalAppointmentDate = JSON.parse(localStorage.getItem("appointment_info"));
  let originalAppointmentDateStartDate = new Date(originalAppointmentDate.start);
  let originalAppointmentDateStartEnd = new Date(originalAppointmentDate.end);
  let convertedToLocalStartUtc = Office.context.mailbox.convertToLocalClientTime(originalAppointmentDateStartDate);
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

  Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
    console.log(`getAllAsync(): asyncResult.status = ${asyncResult.status}`);
    if (asyncResult.status != "failed") {
      for (let index = 0; index < asyncResult.value.length; index++) {
        const element = asyncResult.value[index];
        if (element.key === "timeChanged") {
          //Only need to set the message once
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
      function (asyncResult) {
        console.log("replaceAsync() for 'timeChanged' completed");
        console.dir(asyncResult);
        dateStampMessageSet = true;
        event.completed();
      }
    );
  });
}

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
Office.actions.associate("onMessageAttachmentsChangedHandler", onItemAttachmentsChangedHandler);
Office.actions.associate("onAppointmentAttachmentsChangedHandler", onItemAttachmentsChangedHandler);
Office.actions.associate("onAppointmentAttendeesChangedHandler", onAppointmentAttendeesChangedHandler);
Office.actions.associate("onAppointmentTimeChangedHandler", onAppointmentTimeChangedHandler);
