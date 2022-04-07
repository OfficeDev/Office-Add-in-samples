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

export async function insertChekhovQuoteAtTheBeginning() {
    await Word.run(async (context) => {

        // Create a proxy object for the document body.
        const body = context.document.body;

        // Queue a command to insert text at the start of the document body.
        body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Added a quote from Anton Chekhov.');
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}

export async function insertChineseProverbAtTheEnd() {
    await Word.run(async (context) => {

        // Create a proxy object for the document body.
        const body = context.document.body;

        // Queue a command to insert text at the end of the document body.
        body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Added a quote from a Chinese proverb.');
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
