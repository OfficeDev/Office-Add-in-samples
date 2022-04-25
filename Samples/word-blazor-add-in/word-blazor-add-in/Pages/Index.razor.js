

// -------------------------------------------
// Step 1: Add some paragraphs to the document
// -------------------------------------------

export async function setupDocument() {

    await Word.run(async (context) => {

        context.document.body.clear();
        context.document.body.insertParagraph("One more paragraph. ", "Start");
        context.document.body.insertParagraph("Co-locating Index.razor.js Demo", "Start");
        context.document.body.insertParagraph("Inserting another paragraph. ", "Start");

        context.document.body.insertParagraph(
            "Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance.",
            "Start"
        );

        context.document.body.paragraphs
            .getLast()
            .insertText(
                "With Word add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to build a solution that can run in Word across multiple platforms, including on the web, Windows, Mac, and iPad. Learn how to build, test, debug, and publish Word add-ins.",
                "Replace"
            );
    });
}

// ---------------------------------------------------
// Step 2: Create content controls from the paragraphs
// ---------------------------------------------------

async function handleContentControlAdded(args) {
    console.log("Content control added!");
}

export async function insertContentControls() {

    // Traverses each paragraph of the document and wraps a content control on each with either a even or odd tags.
    await Word.run(async (context) => {

        const currentdocument = context.document;
        currentdocument.load("$all");

        await context.sync();

        currentdocument.onContentControlAdded.add(handleContentControlAdded);

        let paragraphs = context.document.body.paragraphs;
        paragraphs.load("$none"); // Don't need any properties; just wrap each paragraph with a content control.

        await context.sync();

        let contentcontrolsinserted = 0;

        for (let i = 0; i < paragraphs.items.length; i++) {
            let contentControl = paragraphs.items[i].insertContentControl();
            contentcontrolsinserted++;
        }

        await context.sync();
    });
}


// -----------------------------------------------------------------
// Step 3: Tag each content control, by marking them as even and odd
// -----------------------------------------------------------------

async function handleContentControlDeleted(args) {
    console.log("Content control deleted!");
}

async function handleSelectionChanged(args) {
    console.log("Content control selection changed!");
}

export async function tagContentControls() {

    // Traverses each content control of the document and wraps a content control on each with either a even or odd tags.
    await Word.run(async (context) => {

        const currentdocument = context.document;
        currentdocument.load("$all");

        await context.sync();

        let contentcontrolsTagged = 0;

        let contentcontrols = currentdocument.contentControls;
        context.load(contentcontrols);
        await context.sync();

        for (let i = 0; i < contentcontrols.items.length; i++) {
            let contentControl = contentcontrols.items[i];
            // For even, tag "even".
            if (i % 2 === 0) {
                // Tag
                contentControl.tag = "even";
                console.log("Content control tagged even!");
            } else {
                contentControl.tag = "odd";
                console.log("Content control tagged odd!");
            }

            contentcontrolsTagged++;
        }

        await context.sync();
        console.log("Content controls tagged and handled: " + contentcontrolsTagged);
    });
}

// -----------------------------------------------------------------
// Step 4: Modify the content controls to show off the change options
// -----------------------------------------------------------------

export async function modifyContentControls() {

    // Adds title and colors to odd and even content controls and changes their appearance.
    await Word.run(async (context) => {

        // Gets the complete sentence (as range) associated with the insertion point.
        let evenContentControls = context.document.contentControls.getByTag("even");
        let oddContentControls = context.document.contentControls.getByTag("odd");

        evenContentControls.load("length");
        oddContentControls.load("length");

        await context.sync();

        for (let i = 0; i < evenContentControls.items.length; i++) {
            // Change a few properties and append a paragraph.
            evenContentControls.items[i].set({
                color: "red",
                title: "Even ContentControl #" + (i + 1),
                appearance: "Tags"
            });
            evenContentControls.items[i].insertHtml("This is an <strong>even</strong> content control", "End");
        }

        for (let j = 0; j < oddContentControls.items.length; j++) {
            // Change a few properties and append a paragraph.
            oddContentControls.items[j].set({
                color: "green",
                title: "Odd ContentControl #" + (j + 1),
                appearance: "Tags"
            });
            oddContentControls.items[j].insertParagraph("This is an odd content control", "End");
        }

        await context.sync();
    });
}

// --------------------------------------------------------------------------------
// Step 5: Register the content controls for onDeleted and onSelectionChanged events
// --------------------------------------------------------------------------------
let eventContexts = [];

export async function registerEvents() {
    // Traverses each content control of the document and deletes the even content controls
    await Word.run(async (context) => {

        let contentcontrols = context.document.contentControls;
        contentcontrols.load("items");
        await context.sync();

        if (contentcontrols.items.length === 0) {
            console.log("There aren't any content controls in this document so can't register event handlers.");
        } else {
            for (let i = 0; i < contentcontrols.items.length; i++) {
                eventContexts[i*2] = contentcontrols.items[i].onDeleted.add(handleContentControlDeleted);
                console.log("Added onDeleted handler.");
                eventContexts[(i * 2) + 1] = contentcontrols.items[i].onSelectionChanged.add(handleSelectionChanged);
                console.log("Added onSelectionChanged handler.");
                contentcontrols.items[i].track();
            }

            await context.sync();

            console.log("Added onDeleted and onSelectionChanged event handlers.");
        }
    });
}

// -----------------------------------------------------------------------------------
// Step 6: Deregister the content controls for onDeleted and onSelectionChanged events
// -----------------------------------------------------------------------------------

export async function deregisterEvents() {
    await Word.run(async (context) => {
        for (let i = 0; i < eventContexts.length; i++) {
            await Word.run(eventContexts[i].context, async (context) => {
                eventContexts[i].remove();
                console.log("Remove context " + i);
            });
        }

        await context.sync();

        eventContexts = null;
        console.log("Remove the onDeleted and onSelectionChanged event handlers.");
    });
}

// -------------------------------------------
// Step 7: Delete first 'even' content control
// -------------------------------------------

export async function deleteContentControl() {
    await Word.run(async (context) => {
        let contentControls = context.document.contentControls.getByTag("even");
        contentControls.load("items");
        await context.sync();

        if (contentControls.items.length === 0) {
            console.log("There are no content controls tagged 'even' in this document.");
        } else {
            console.log("First 'even' control to be deleted:");
            console.log(contentControls.items[0]);
            contentControls.items[0].delete(false);
            await context.sync();
        }
    });
}
