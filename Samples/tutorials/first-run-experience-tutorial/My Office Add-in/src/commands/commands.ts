/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Runs an add-in command function.
 * To learn more about function commands, see
 * https://learn.microsoft.com/office/dev/add-ins/design/add-in-commands#types-of-add-in-commands.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // Your code goes here.

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
