/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

/**
 * Clears the document
 */
export async function clearDocument() {
    await Word.run(async (context) => {
        context.document.body.clear();
    });
}
/**
 * Inserts Paragraphs on the specified location
 * @param  {} text
 * @param  {} location
 */
export async function insertParagraph(text, location) {
    await Word.run(async (context) => {
        context.document.body.insertParagraph(text, location);
    });
}
/**
 * Replacing text in the last paragraph found in the document
 * @param  {} text
 */
export async function replaceParagraph(text) {
    await Word.run(async (context) => {
        context.document.body.paragraphs
            .getLast()
            .insertText(
                text,
                "Replace"
            );
    });
}
/**
 * This will count the number of paragraphs in the document
 * @returns numberofParagraphs
 */
export async function paragraphCount() {

        let numberofParagraphs = 0;

        await Word.run(async (context) => {
            const currentdocument = context.document;
            currentdocument.load("$all");

            await context.sync();

            let paragraphs = context.document.body.paragraphs;
            paragraphs.load("$none"); // Don't need any properties;

            await context.sync();

            numberofParagraphs = paragraphs.items.length;

            console.log("Paragraph Count JS: ");
            console.log(numberofParagraphs);
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    
    return { Value: numberofParagraphs } ;
}
