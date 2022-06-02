/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
/**
 * Inserts a quote by Emerson into the current cursor location.
 */
export async function insertEmersonQuoteAtSelection() {
    await Word.run(async (context) => {

        // Create a proxy object for the document.
        const thisDocument = context.document;

        // Queue a command to get the current selection.
        // Create a proxy range object for the selection.
        const range = thisDocument.getSelection();

        // Queue a command to replace the selected text.
        range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Added a quote from Ralph Waldo Emerson.');
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
/**
 * Inserts a quote by Chekhov at the start of the document body.
 */
export async function insertBruceSchneierQuoteAtTheBeginning() {
    await Word.run(async (context) => {

        // Create a proxy object for the document body.
        const body = context.document.body;

        // Queue a command to insert text at the start of the document body.
        body.insertText('"There is an entire flight simulator hidden in every copy of Microsoft Excel 97."\n', Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Added a quote from Bruce Schneier.');
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
/**
 * Inserts a quote from Steve Ballmer at the end of the document.
 */
export async function insertSteveBallmerAtTheEnd() {
    await Word.run(async (context) => {

        // Create a proxy object for the document body.
        const body = context.document.body;

        // Queue a command to insert text at the end of the document body.
        body.insertText('"Developer, developer, developer!"\n', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Added a quote from Steve Ballmer.');
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
