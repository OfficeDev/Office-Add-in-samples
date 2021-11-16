/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
*/

var fileName;
const secretKey = "secret key 123";
const encryptedAttachmentPrefix = "encrypted_";
const decryptedAttachmentPrefix = "decrypted_";
var totalOptionalAttendees = 0;
var totalRequiredAttendees = 0;
var totalDistributionLists = 0;
// eslint-disable-next-line no-unused-vars
Office.initialize = function (reason) {
  try {
    console.log(`Office.initialize(): Huzzah!`);
  }
  catch(ex){
    console.error(`Office.initialize(): Error! ${ex}`);
  }  
};
/**
 * Method that fires when a appointment is being created or edited
 * @param {Office.AsyncResult} event default: Office.AsyncResult
 */
function onAppointmentComposeHandler(event) {
  //NOTE: Must call event.completed() when any logic below is finished!

  console.log("onAppointmentComposeHandler(): entered!");

  let originalAppointmentDate = {}; //Create an object to cache the original date/time and persist it to localStorage

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
          //Add a notification message to ask the user to open the Task Pane to view additional information on the sample
          //NOTE: Clicking the "Show Task Pane" link in the InfoBar doesn't work in Outlook Online. A fix is in progress and being tested: https://github.com/OfficeDev/office-js/issues/2125
                      
          console.log("onAppointmentComposeHandler(): Adding notification message...");

          //NOTE: actions array only applicable to insightMessage types
          Office.context.mailbox.item.notificationMessages.addAsync(
            "showInfoBarForSampleInstructions",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
              message:
                "Open the Task Pane for details about running the Outlook Event-based Activation Sample Add-in",
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
 * Method that fires when an email is being created 
 * @param {Office.AsyncResult} result default: Office.AsyncResult
 */
function onMessageComposeHandler(event) {
  console.log("onMessageComposeHandler(): entered!");

  //Displays the InfoBar in the compose email. Purpose is to remind the user running this sample to open the Task Pane to display additional instructions and reference information
  //NOTE: Clicking the "Show Task Pane" link in the InfoBar doesn't work in Outlook Online. A fix is in progress and being tested: https://github.com/OfficeDev/office-js/issues/2125
  //NOTE: actions array only applicable to insightMessage types
  var options = { asyncContext: { callingEvent: event } };
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
 * Method that fires when an attendee is added or removed from an appointment
 * @param {Office.AsyncResult} event default: Office.AsyncResult
 */
function onAppointmentAttendeesChangedHandler(event) {
  
  //NOTE: Must call event.completed() when any logic below is finished!

  var totalOptionalAttendees = 0;
  var totalRequiredAttendees = 0;
  var totalDistributionLists = 0;

  console.log(`onAppointmentAttendeesChangedHandler(): type = ${event.type}; requiredAttendees: ${event.changedRecipientFields.requiredAttendees}; optionalAttendees: ${event.changedRecipientFields.optionalAttendees}; resources: ${event.changedRecipientFields.resources};`);

  //Run a series of Office async calls to calculate the total number of current required and optional attendees as well as distribution lists. Then add, update or remove notification messages depending on the counts. These calls are carried out sequentially in nested callbacks but could be refactored to use closures or Promises for cleaner code.

  //Get required attendees
  Office.context.mailbox.item.requiredAttendees.getAsync(function (asyncResult) {
    console.log(`onAppointmentAttendeesChangedHandler(): getAsync => asyncResult.status: ${asyncResult.status}`);  
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      //Add to our total of required attendees and ensure that the current user isn't processed as an attendee
      var apptRequiredAttendees = asyncResult.value;

      for (var i = 0; i < apptRequiredAttendees.length; i++) {        
        console.log(
          `onAppointmentAttendeesChangedHandler(): Required attendee: ${apptRequiredAttendees[i].displayName} (${apptRequiredAttendees[i].emailAddress}) - type: ${apptRequiredAttendees[i].recipientType}`
        );
      }

      //Get attendees except for the current user (which is only included in this array on Outlook on Windows, not on Outlook on the web.)
      console.log(`onAppointmentAttendeesChangedHandler(): Filtering out attendees for current user ${Office.context.mailbox.userProfile.emailAddress}...`);
      var apptRequiredAttendeesWithoutUser = apptRequiredAttendees.filter(function(attendee){return attendee.emailAddress !== Office.context.mailbox.userProfile.emailAddress;});
      
      totalRequiredAttendees = apptRequiredAttendeesWithoutUser.length;
      currentDistributionLists = apptRequiredAttendees.filter(function(attendee){return attendee.recipientType === "distributionList";});
      if (currentDistributionLists.length !== 0){
        totalDistributionLists = totalDistributionLists + currentDistributionLists.length;
      }     

      console.log(`onAppointmentAttendeesChangedHandler(): totalRequiredAttendees = ${totalRequiredAttendees}; totalDistributionLists = ${totalDistributionLists}`);

      //Get optional attendees
      Office.context.mailbox.item.optionalAttendees.getAsync(function (asyncResult2) {
        console.log(`onAppointmentAttendeesChangedHandler(): getAsync => asyncResult2.status: ${asyncResult2.status}`);          
        if (asyncResult2.status === Office.AsyncResultStatus.Succeeded) {
          //Add to our total of optional attendees and ensure that the current user isn't processed as an attendee
          var apptOptionalAttendees = asyncResult2.value;          
          //Get attendees except for the current user (which is somehow only included in this array on Outlook Desktop, not Outlook Online)
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

        //Update notification messages
        //=============================================================================================
        if (totalOptionalAttendees === 0 && totalRequiredAttendees === 0) {
          //Remove the info bar with the recipients tally and any distribution list warnings if there are no longer any recipients
          Office.context.mailbox.item.notificationMessages.removeAsync(
            "attendeesChanged",
            null,
            function (asyncResult3) {
              console.log(`onAppointmentAttendeesChangedHandler() removeAsync:attendeesChanged => asyncResult3.status: ${asyncResult3.status}`);     
              //No need to call if there are no more distribution lists; otherwise, remove the warning about distribution lists
              if (asyncResult3.status === Office.AsyncResultStatus.Succeeded && totalDistributionLists !== 0) {
                Office.context.mailbox.item.notificationMessages.removeAsync(
                  "distributionListWarning",
                  null,
                  function (asyncResultDLs) {
                    console.log(`onAppointmentAttendeesChangedHandler() removeAsync:distributionListWarning => asyncResultDLs.status: ${asyncResultDLs.status}`);     
                    event.completed(); //NOTE: Must call!
                  }
                );
              } else {
                //No more attendees
                event.completed(); //NOTE: Must call!
              }
            }
          );
        } else 
        {
          //Update the info bar with the recipients tally
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

                  //TEST Only call notificationMessages.getAllAsync if totalDistributionLists === 0?
                  if (totalDistributionLists === 0 && anyExistingWarningMessages === true) {        
                    //Remove the warning about distribution lists            
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

                    //Update existing warning about distribution lists
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
                console.error(`onAppointmentAttendeesChangedHandler(): Error with item.notificationMessages.replaceAsync(): ${asyncResult4.error}`);
              }
            }
          );
        }
        //=============================================================================================
      });
    } else {
      console.error(`onAppointmentAttendeesChangedHandler(): Unexpected: asyncResult.status = ${asyncResult.status}`);
      event.completed(); //NOTE: Must call!
    }
  });
}
/**
 * Method that fires when an attachment is being added or removed from the compose email or appointment
 * @param {Office.AsyncResult} event default: Office.AsyncResult
 */
function onItemAttachmentsChangedHandler(event) {  
  if (Office.context.platform  !== Office.PlatformType.OfficeOnline){
    console.warn(`onItemAttachmentsChangedHandler(): Unsupported platform for encrypting/decrypting attachments (${Office.context.platform}); leaving...`);
    event.completed();
    return;
  }

  if (event.attachmentStatus === "removed"){
    console.log("onItemAttachmentsChangedHandler(): Not processing removed attachments; leaving...");
    event.completed(); //NOTE: Must call!
    return;
  }

  console.log(`onItemAttachmentsChangedHandler(): ${event.attachmentDetails.name} (${event.attachmentStatus})`);
  
    if (event.attachmentDetails.name == `${decryptedAttachmentPrefix}${fileName}`) {
    //Don't process any more events - we've already encrypted the attachment and added it as another attachment, then decrypted that attachment and added it as well
    event.completed(); //NOTE: Must call!
    return;
  }
  if (fileName !== undefined){
    console.log("onItemAttachmentsChangedHandler(): Skipping processing of further attachments - demo is done!");
    event.completed(); //NOTE: Must call!
    return;
  }

  //Process the first attachment. We'll encrypt it and add it as another attachment, then decrypt that attachment and add it as well
  fileName = event.attachmentDetails.name;
  var item = Office.context.mailbox.item;
  var options = { asyncContext: { currentItem: item, callingEvent: event } };    
  item.getAttachmentsAsync(options, getAttachmentsCallback);
}
/**
 * Processes the attachments in the current item. NOTE: Only the first attachment that's added will be processed
 * @param {Office.AsyncResult} result default: Office.AsyncResult
 */
function getAttachmentsCallback(result) { 
  var options = { asyncContext: { callingEvent: result.asyncContext.callingEvent } };    
  //Only handle the first attachment (0 index in the array) - ignore the others
  result.asyncContext.currentItem.getAttachmentContentAsync(result.value[0].id, options, handleAttachmentsCallback);
}
/**
 * Method that encrypts base64 file data using CryptoJS and attaches the file to the email. Cloud, .eml and .ICalendar attachments will not be processed.
 * @param {Office.AsyncResult} result default: Office.AsyncResult
 */
function handleAttachmentsCallback(result) {

  console.log(`handleAttachmentsCallback(): result.value.format = ${result.value.format}`);
  // console.dir(result.value.content); //NOTE: If you want to see the base64 data output to the console, uncomment this line - but console.dir() functions cannot be used when runtime logging is enabled!!

  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      //Handle file attachment      
      //Set a notification message that we're processing the attachment. Note that this will be removed immediately after the decrypted attachment is added, and it may not be displayed for very long
      var options = { 'asyncContext': { base64: result.value.content, callingEvent: result.asyncContext.callingEvent } };
      Office.context.mailbox.item.notificationMessages.addAsync("processingAttachments", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
        message: `Please wait while the '${fileName}' attachment is encrypted...`,
      }, options, function(asyncResult){
          //Encrypt base64 file data using CryptoJS
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            try {
              var ciphertext = CryptoJS.AES.encrypt(asyncResult.asyncContext.base64, secretKey).toString();
              //Then attaches the file to the email
              console.log(`handleAttachmentsCallback(): starting processing of file '${fileName}'...`);
              encryptAttachment(ciphertext, asyncResult.asyncContext.callingEvent);
            } catch (ex) {
              console.error(`handleAttachmentsCallback(): Error: ${ex}`);
              options = { 'asyncContext': { callingEvent: asyncResult.asyncContext.callingEvent } };
              Office.context.mailbox.item.notificationMessages.removeAsync("processingAttachments", options, function (asyncResult2) {
                console.log("handleAttachmentsCallback(): Notification message removed.");
                asyncResult2.asyncContext.callingEvent.completed();
              });              
            }
          } else {
            console.error(`handleAttachmentsCallback(): Unexpected - status is ${asyncResult.status}`);
            asyncResult.asyncContext.callingEvent.completed();
          }     
      });      
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      console.log("handleAttachmentsCallback(): Attachment is a message.");
      result.asyncContext.callingEvent.completed();
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      console.log("handleAttachmentsCallback(): Attachment is a calendar item.");
      result.asyncContext.callingEvent.completed();
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      console.log("handleAttachmentsCallback(): Attachment is a cloud attachment.");
      result.asyncContext.callingEvent.completed();
      break;
    default:
      // Handle attachment formats that are not supported.
      console.console.warn();("handleAttachmentsCallback(): Not handling unsupported attachment.");
      result.asyncContext.callingEvent.completed();
      break;
  }
}
/**
 * Method that converts encrypted data to base64 and creates and adds a file attachment to the current email
 * @param {string} encryptedData default: "undefined"
 * @param {Office.AsyncResult} callingEvent default: Office.AsyncResult. The event parameter from the action event
 */
function encryptAttachment(encryptedData, callingEvent) {
  console.log(`encryptAttachment(): Encrypting file '${fileName}'...`);
  // console.dir(encryptedData); //NOTE: If you want to see the encrypted data output to the console, uncomment this line

  var base64EncryptedData = window.btoa(encryptedData);
  var encryptedFileName = `${encryptedAttachmentPrefix}${fileName}`;
  var options = { 'asyncContext': { encryptedFileName: encryptedFileName, callingEvent: callingEvent}, 'isInline': false };

  //NOTE: If you want to see the base64 data output to the console, uncomment these lines
  // console.log("encryptAttachment(): base64 encrypted data:");
  // console.dir(base64EncryptedData);

  console.log(`encryptAttachment(): Adding encrypted file '${encryptedFileName}'...`);
  Office.context.mailbox.item.addFileAttachmentFromBase64Async(
    base64EncryptedData,
    encryptedFileName,
    options,
    function (asyncResult) {
      options = { 'asyncContext': { encryptedFileName: asyncResult.asyncContext.encryptedFileName, callingEvent: asyncResult.asyncContext.callingEvent, encryptedData: encryptedData} };
      console.log(`encryptAttachment(): Added encrypted attachment '${asyncResult.asyncContext.encryptedFileName}'; now decrypting...`);
      //console.dir(asyncResult); //NOTE: If you want to see the base64 data output to the console, uncomment this line
      decryptAttachment(options);
    }
  );
}
/**
 * Method that decrypts encrypted base64 file data using CryptoJS and attaches the file to the email
 * @param {Office.AsyncResult} result default: Office.AsyncResult. Object containing encryptedFileName, encryptedData and callingEvent properties in asyncContext property
 */
function decryptAttachment(result) {
  console.log(`decryptAttachment(): Decrypting file '${result.asyncContext.encryptedFileName}'...`);
  var bytes = CryptoJS.AES.decrypt(result.asyncContext.encryptedData, secretKey);
  var originalText = bytes.toString(CryptoJS.enc.Utf8);
  var decryptedFileName = `${decryptedAttachmentPrefix}${fileName}`;

  // console.log(`decryptAttachment(): Original base64: ${originalText}`); //NOTE: If you want to see the base64 data output to the console, uncomment this line
  console.log(`decryptAttachment(): Adding decrypted file '${decryptedFileName}'...`);

  var options = { 'asyncContext': { callingEvent: result.asyncContext.callingEvent} };
  Office.context.mailbox.item.addFileAttachmentFromBase64Async(
    originalText,
    decryptedFileName,
    options,
    function (asyncResult) {
      console.log(`decryptAttachment(): Added decrypted attachment '${decryptedFileName}'`);
      // console.dir(asyncResult); //NOTE: If you want to see the base64 data output to the console, uncomment this line
      Office.context.mailbox.item.notificationMessages.removeAsync("processingAttachments", options, function (asyncResult2) {
        console.log("decryptAttachment(): Notification message removed.");        
        Office.context.mailbox.item.notificationMessages.addAsync("attachmentsAdded", {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: `The '${fileName}' attachment has been encrypted and decrypted and added as reference attachments for your review.`,
          icon: "Icon.16x16",
          persistent: false,
        }, options, function(asyncResult3){
          console.log("decryptAttachment(): All attachment processing operations completed!");        
          asyncResult3.asyncContext.callingEvent.completed();
        });
      });
    }
  );
}

/**
 * Method that fires when the user changes the date/time for an appointment
 * @param {Office.AsyncResul} result default: Office.AsyncResult
 */
function onAppointmentTimeChangedHandler(event) {
  console.log(`onAppointmentTimeChangedHandler(): type: ${event.type}; start: ${event.start}; end: ${event.end}`);

  Office.context.mailbox.item.sessionData.getAsync("appointment_info", function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(`onAppointmentTimeChangedHandler(): sessionData retrieved: ${asyncResult.value}`);

      try {
        //Convert the start and end dates to appropriate formats
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

        //Add a notification message with the original start/end date/time (if it hasn't already been added)
        Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult2) {
          console.log(`getAllAsync(): asyncResult.status = ${asyncResult2.status}`);
          if (asyncResult2.status != "failed") {
            for (let index = 0; index < asyncResult2.value.length; index++) {
              const element = asyncResult2.value[index];
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

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
Office.actions.associate("onMessageAttachmentsChangedHandler", onItemAttachmentsChangedHandler);
Office.actions.associate("onAppointmentAttachmentsChangedHandler", onItemAttachmentsChangedHandler);  
Office.actions.associate("onAppointmentAttendeesChangedHandler", onAppointmentAttendeesChangedHandler);
Office.actions.associate("onAppointmentTimeChangedHandler", onAppointmentTimeChangedHandler);
