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
 * Inserts text at a given location (here hard-coded 😱)
 * @param  {} text
 * @param  {} location
 */
export async function requestContextDemo(text, location) {
    var ctx = new Word.RequestContext();
    var range = ctx.document.getSelection();

    range.insertText("Test MinimalWordMethod", "After");

    await ctx.sync();
}
