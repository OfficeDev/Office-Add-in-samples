/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

Office.onReady();

/**
 * The words in the subject or body that require corresponding color categories to be applied to a new
 * message or appointment.
 * @constant
 * @type {string[]}
 */
 const KEYWORDS = ["sales", "expense reports", "legal", "marketing", "performance reviews"];

 /**
  * Handle the OnNewMessageCompose or OnNewAppointmentOrganizer event by verifying that keywords have corresponding 
  * color categories when a new message or appointment is created. If no corresponding categories exist, they will be
  * created.
  * @param {Office.AddinCommands.Event} event The OnNewMessageCompose or OnNewAppointmentOrganizer event object.
  */
 function onItemComposeHandler(event) {
   Office.context.mailbox.masterCategories.getAsync({ asyncContext: event }, asyncResult => {
     let event = asyncResult.asyncContext;
 
     if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
       let categories = asyncResult.value;
       let categoriesToBeCreated = [];
       if (categories) {
         let categoryNamesInUse = getCategoryProperty(categories, "displayName");
         let categoryColorsInUse = getCategoryProperty(categories, "color");
         categoriesToBeCreated = getCategoriesToBeCreated(KEYWORDS, categoryNamesInUse);
         if (categoriesToBeCreated.length > 0) {
           categoriesToBeCreated = assignCategoryColors(categoriesToBeCreated, categoryColorsInUse);
         }
       } else {
         categoriesToBeCreated = assignCategoryColors(getCategoriesToBeCreated(KEYWORDS));
       }
       
       createCategories(event, categoriesToBeCreated);
       event.completed({ allowEvent: true });
     } else {
       console.log(asyncResult.error.message);
       event.completed({ allowEvent:false, errorMessage: "Failed to configure categories." });
       return;
     }
   });
 }
 
 /**
  * Handle the OnMessageSend or OnAppointmentSend event by verifying that applicable color categories are 
  * applied to a new message or appointment before it's sent.
  * @param {Office.AddinCommands.Event} event The OnMessageSend or OnAppointmentSend event object.
  */
 function onItemSendHandler(event) {
   Office.context.mailbox.item.subject.getAsync({ asyncContext: event }, asyncResult => {
     let event = asyncResult.asyncContext;
 
     if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
       let subject = asyncResult.value;
       let detectedWords = [];
       if (subject) {
           detectedWords = checkForKeywords(KEYWORDS, subject);
       }
       
       let options = { asyncContext: { callingEvent: event, keywordArray: detectedWords }};
       Office.context.mailbox.item.body.getAsync("text", options, asyncResult => {
         let event = asyncResult.asyncContext.callingEvent;
           
         if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
           let body = asyncResult.value;
           let detectedWords = asyncResult.asyncContext.keywordArray;
           if (body) {
             detectedWords = checkForKeywords(KEYWORDS, body, detectedWords);
           }
 
           if (detectedWords.length > 0) {
             checkAppliedCategories(event, detectedWords);
           } else {
             event.completed({ allowEvent: true });
           }
         } else {
             console.log(asyncResult.error.message);
             event.completed({ allowEvent: false, errorMessage: "Failed to check the body for keywords." });
             return;
         }
       });
     } else {
         console.log(asyncResult.error.message);
         event.completed({ allowEvent: false, errorMessage: "Failed to check the subject for keywords." });
         return;
     }
   });
 }
 
 /**
  * Get the property values of existing categories.
  * @param {Office.CategoryDetails[]} categories Existing categories in Outlook.
  * @param {string} property The property to extract from existing categories. Categories have a display name and a color.
  * @returns {string[]} The property's value.
  */
  function getCategoryProperty(categories, property) {
   let values = [];
   categories.forEach(category => {
     values.push(category[property]);
   });
 
   return values;
 }
 
 /**
  * Determine the categories to be created based on existing categories.
  * @param {string[]} keywords The keywords that require corresponding categories.
  * @param {string[]} existingCategories The display names currently in use by existing categories.
  * @returns {string[]} The names of the new categories.
  */
 function getCategoriesToBeCreated(keywords, existingCategories = []) {
   let categoriesToBeCreated = [];
   if (existingCategories.length === 0) {
     keywords.forEach(word => {
       categoriesToBeCreated.push(`Office Add-ins Sample: ${word}`);
     });
   } else {
     keywords.forEach(word => {
       if (!existingCategories.includes(`Office Add-ins Sample: ${word}`)) {
         categoriesToBeCreated.push(`Office Add-ins Sample: ${word}`);
       }
     });
   }
 
   return categoriesToBeCreated;
 }
 
 /**
  * Assign a color to a new category based on available colors. If all 25 colors are in use,
  * duplicate colors are assigned starting from Preset0.
  * @param {string[]} categoriesToBeCreated The names of the new categories.
  * @param {string[]} categoryColorsInUse The colors currently in use by existing categories.
  * @returns {Office.CategoryDetails[]} The new category objects to be created.
  */
 function assignCategoryColors(categoriesToBeCreated = [], categoryColorsInUse = []) {
   const totalColors = 25;
   if (categoryColorsInUse.length >= totalColors) {
     for (let i = 0; i < categoriesToBeCreated.length; i++) {
       categoriesToBeCreated[i] = {
         displayName: categoriesToBeCreated[i],
         color: `Preset${i}`
       };
     }
   } else {
     for (let i = 0; i < categoriesToBeCreated.length; i++) {
       for (let j = 0; j < totalColors; j++) {
         if (!categoryColorsInUse.includes(`Preset${j}`)) {
           categoriesToBeCreated[i] = {
             displayName: categoriesToBeCreated[i],
             color: `Preset${j}`
           };
 
           categoryColorsInUse.push(`Preset${j}`);
           break;
         }
       }
     }
   }
 
   return categoriesToBeCreated;
 }
 
 /**
  * Create categories.
  * @param {Office.AddinCommands.Event} event The OnNewMessageCompose or OnNewAppointmentOrganizer event object.
  * @param {Office.CategoryDetails[]} categoriesToBeCreated The new category objects to create.
  */
 function createCategories(event, categoriesToBeCreated) {
   Office.context.mailbox.masterCategories.addAsync(categoriesToBeCreated, { asyncContext: event }, asyncResult => {
     if (asyncResult.status === Office.AsyncResultStatus.Failed) {
       console.log(asyncResult.error.message);
       asyncResult.asyncContext.completed({ allowEvent: false, errorMessage: "Failed to set new categories." });
       return;
     }
   });
 }
 
 /**
  * Determine if keywords are present in the message or appointment's subject or body that require corresponding categories.
  * @param {string[]} keywords The keywords that require corresponding categories.
  * @param {string} text The contents of the subject or body of the message or appointment.
  * @param {string[]} detectedWords The keywords found in the message or appointment's subject or body.
  * @returns {string[]} Keywords detected in the message or appointment's subject or body that require corresponding categories.
  */
 function checkForKeywords(keywords, text, detectedWords = []) {
   keywords = new RegExp(keywords.join("|"), "gi");
   text = text.toLowerCase();
 
   let keywordsFound = text.match(keywords);
   if (keywordsFound) {
     checkForDuplicates(keywordsFound, detectedWords);
   }
 
   return detectedWords;
 }
 
 /**
  * Check for duplicate keywords in the message or appointment's subject or body.
  * @param {string[]} wordsToCompare The keywords found in the message or appointment's subject or body to compare to the existing
  * list of detected keywords.
  * @param {string[]} wordList The existing list of detected keywords. 
  */
 function checkForDuplicates(wordsToCompare = [], wordList = []) {
   wordsToCompare.forEach(word => {
     if (!wordList.includes(word)) {
      wordList.push(word);
     }
   });
 }
 
 /**
  * Determine the categories to be added based on the detected keywords in the message or appointment's subject or body.
  * @param {string[]} detectedWords The keywords detected in the message or appointment's subject or body.
  * @returns {string[]} The names of the categories to be added to the message or appointment.
  */
 function getCategoryName(detectedWords) {
   let categories = [];
   detectedWords.forEach(word => {
     categories.push(`Office Add-ins Sample: ${word}`);
   });
 
   return categories;
 }
 
 /**
  * Check that the appropriate categories, based on detected keywords in the subject or body, are applied to the 
  * message or appointment before it's sent.
  * @param {Office.AddinCommands.Event} event The OnMessageSend or OnAppointmentSend event object.
  * @param {string[]} detectedWords The keywords found in the message or appointment's subject or body.
  */
  function checkAppliedCategories(event, detectedWords) {
   let options = { asyncContext: { callingEvent: event, keywordArray: detectedWords } };
   Office.context.mailbox.item.categories.getAsync(options, asyncResult => {
     let sendEvent = asyncResult.asyncContext.callingEvent;
     
     if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
       let requiredCategories = getCategoryName(asyncResult.asyncContext.keywordArray);
       let detectedCategories = asyncResult.value;
       if (detectedCategories) {
         let detectedCategoryNames = getCategoryProperty(detectedCategories, "displayName");
         let missingCategories = getMissingCategories(requiredCategories, detectedCategoryNames);
         if (missingCategories.length > 0) {
           let message = `Don't forget to also add the following categories: ${missingCategories.join(", ")}`;
           console.log(message);
           sendEvent.completed({ allowEvent: false, errorMessage: message });
           return;
         }
 
         sendEvent.completed({ allowEvent: true });
       } else {
         let message = `You must assign the following categories before your ${Office.context.mailbox.item.itemType} can be sent: ${requiredCategories.join(", ")}`;
         console.log(message);
         sendEvent.completed({ allowEvent: false, errorMessage: message });
         return;
       }
     } else {
       console.log(asyncResult.error.message);
       sendEvent.completed({ allowEvent: false, errorMessage: "Failed to check categories applied to the item." });
       return;
     }
   });
 }
 
 /**
  * Get the names of the required categories still missing from the message or appointment.
  * @param {string[]} requiredCategories The names of the categories required on the message or appointment before it can be sent.
  * @param {string[]} appliedCategories The names of the categories that are currently applied to the message or appointment.
  * @returns {string[]} The names of the categories that need to be applied to the message or appointment.
  */
 function getMissingCategories(requiredCategories, appliedCategories) {
   let missingCategories = requiredCategories.filter(category => !appliedCategories.includes(category));
   return missingCategories;
 }
 
 Office.actions.associate("onMessageComposeHandler", onItemComposeHandler);
 Office.actions.associate("onAppointmentComposeHandler", onItemComposeHandler);
 Office.actions.associate("onMessageSendHandler", onItemSendHandler);
 Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);