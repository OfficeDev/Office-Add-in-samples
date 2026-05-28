/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office PowerPoint console */

/**
 * Inserts a text box with a colored fill on the selected slide.
 * @param event The add-in command event.
 * @param color The fill color to apply to the text box (CSS color name or hex string). Defaults to "white".
 */
export async function insertTextInPowerPoint(event: Office.AddinCommands.Event, color: string = "white") {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox = slide.shapes.addTextBox(`Hello World (${color})`);
      textBox.fill.setSolidColor(color);
      textBox.lineFormat.color = "black";
      textBox.lineFormat.weight = 1;
      textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
