
export async function clearDocument() {
    await Word.run(async (context) => {
        context.document.body.clear();
    });
}

export async function insertParagraph(text, location) {

     
    await Word.run(async (context) => {
        context.document.body.insertParagraph(text, location);
    });
}

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

            console.log("Paragraph count JS: ");
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
