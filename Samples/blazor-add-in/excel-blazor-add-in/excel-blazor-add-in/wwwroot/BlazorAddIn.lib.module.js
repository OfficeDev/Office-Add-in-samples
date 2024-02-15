/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
/**
 * JavaScript Initializers
 * JavaScript (JS) initializers execute logic before and after a Blazor app loads. JS initializers are useful in the following scenarios:
 * - Customizing how a Blazor app loads.
 * - Initializing libraries before Blazor starts up.
 * - Configuring Blazor settings.
 */

/**
 * beforeStart(options, extensions): 
 * 
 * Called before Blazor starts. 
 * For example, beforeStart is used to customize the loading process, logging level, and other options specific to the hosting model. 
 * @param  {} wasmoptions
 * @param  {} extensions
 */
export async function beforeStart(wasmoptions, extensions) {
    console.log("We are now entering function: beforeStart");

    Office.onReady((info) => {
        // Check that we loaded into Excel.
        if (info.host === Office.HostType.Excel) {
            console.log("We are now hosting in Excel.");
        }
        else {
            console.log("We are now hosting in The Browser (of your choice).");
        }
        console.log("Office onReady.");
    });
}

/**
 * afterStarted: Called after Blazor is ready to receive calls from JS. 
 * For example, afterStarted is used to initialize libraries by making JS interop calls and registering custom elements.
 * The Blazor instance is passed to afterStarted as an argument.
 * @param  {} blazor
 */
export async function afterStarted(blazor) {
    console.log("We are now entering function: afterStarted");
}

/**
 * Writes the event source id to the document when ExecuteFunction runs.
 * @param event {Office.AddinCommands.Event}
 */

function writeValue2(event) {

    console.log("We are now entering function: writeValue");

    Office.context.document.setSelectedDataAsync(
        "ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // Show error message.
            } else {
                // Show success message.
            }
        }
    );

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

Office.actions.associate("writeValue2", writeValue2);