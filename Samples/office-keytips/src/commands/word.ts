/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Word console */

/**
 * When the add-in command is selected, inserts a paragraph with colored text in Word.
 * @param event The add-in command event.
 * @param color The font color to apply to the inserted paragraph (CSS color name or hex string). Defaults to "blue".
 */
export async function insertBlueParagraphInWord(event: Office.AddinCommands.Event, color: string = "blue") {
  try {
    await Word.run(async (context) => {
      const paragraph = context.document.body.insertParagraph(`Hello World (${color})`, Word.InsertLocation.end);
      paragraph.font.color = color;
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
