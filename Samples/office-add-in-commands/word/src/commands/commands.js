/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window, Word */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Writes the event source ID to the document when ExecuteFunction runs.
 * @param event {Office.AddinCommands.Event}
 */

 async function writeValue(event) {
  Word.run(async (context) => {
    // Insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(
      "ExecuteFunction works. Button ID=" + event.source.id,
      Word.InsertLocation.end
    );

    // Change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

Office.actions.associate("writeValue", writeValue);