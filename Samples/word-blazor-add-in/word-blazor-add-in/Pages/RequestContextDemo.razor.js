
export async function clearDocument() {
    await Word.run(async (context) => {
        context.document.body.clear();
    });
}

export async function requestContextDemo(text, location) {
    var ctx = new Word.RequestContext();
    var range = ctx.document.getSelection();

    range.insertText("Test MinimalWordMethod", "After");

    await ctx.sync();
}
