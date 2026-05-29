/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Excel console */

/**
 * When the add-in command is selected, sets the selected range's fill color in Excel.
 * @param event The add-in command event.
 * @param color The fill color to apply (CSS color name or hex string). Defaults to "yellow".
 */
export async function setRangeColorInExcel(event: Office.AddinCommands.Event, color: string = "yellow") {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = color;
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
