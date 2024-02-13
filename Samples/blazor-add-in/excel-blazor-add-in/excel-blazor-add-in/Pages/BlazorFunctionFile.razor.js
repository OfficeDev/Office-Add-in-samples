console.log("Loading BlazorFunctionFile.js");

// The command function.
async function highlightSelection(event) {

    // Implement your custom code here. The following code is a simple Excel example.  
    try {
        console.log("Running highlightSelection");

        console.log("Before callStaticLocalComponentMethodinit");
        callStaticLocalComponentMethodinit();
        console.log("After callStaticLocalComponentMethodinit");

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "red";
            await context.sync();
        });

    } catch (error) {
        // Note: In a production add-in, notify the user through your add-in's UI.
        console.error(error);
    }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

async function callStaticLocalComponentMethodinit() {
    console.log("In callStaticLocalComponentMethodinit");
    try {
        var name = "init";
        // Call JSInvokable Function here ...

        name = await DotNet.invokeMethodAsync("BlazorAddIn", "SayHello", "Skod from Blazor");
        console.log("fin init : " + name)

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.values = [[name]];
            await context.sync();
        });

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "yellow";
            await context.sync();
        });
    }
    catch (err) {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.values = [[err.message]];
            await context.sync();
        });
        console.log();
        console.log("Error call : " + err.message);
    }
    finally {
        console.log("Finish callStaticLocalComponentMethodinit");
    }
}

// You must register the function with the following line.
Office.actions.associate("highlightSelection", highlightSelection);