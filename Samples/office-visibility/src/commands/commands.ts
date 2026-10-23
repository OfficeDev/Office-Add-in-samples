/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office console */

const commandActions = [
  ["button1Action", "Button 1"],
  ["button2Action", "Button 2"],
  ["secondaryOneAction", "Secondary button one"],
  ["secondaryTwoAction", "Secondary button two"],
  ["menuItemOneAction", "Menu item one"],
  ["menuItemTwoAction", "Menu item two"],
] as const;

function createCommandHandler(label: string) {
  return (event: Office.AddinCommands.Event) => {
    console.log(`Ribbon command selected: ${label}`);
    event.completed();
  };
}

Office.onReady(() => {
  for (const [actionId, label] of commandActions) {
    Office.actions.associate(actionId, createCommandHandler(label));
  }
});
