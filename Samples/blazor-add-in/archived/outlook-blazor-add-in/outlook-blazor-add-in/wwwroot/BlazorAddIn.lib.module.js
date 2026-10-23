/* Copyright(c) Maarten van Stam and Eric Legault. All rights reserved. Licensed under the MIT License. */
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
    //This does NOT log when hosted in full browser instance outside of the Outlook Task pane
    console.log("beforeStart: BlazorAddIn.lib.module.js()");
    console.log("beforeStart: Entering function");

    Office.onReady((info) => {
        console.log("beforeStart: Entering Office.onReady ...");

        // Check that we loaded into Outlook, what is the Office Host here
        if (info.host === Office.HostType.Outlook) {
            console.log("beforeStart: We are now hosting in Outlook");
        }
        else {
            console.log("beforeStart: We are now hosting in The Browser (of your choice)");
        }

        console.log("beforeStart: Finished Office onReady");
    });
}

/**
 * afterStarted: Called after Blazor is ready to receive calls from JS. 
 * For example, afterStarted is used to initialize libraries by making JS interop calls and registering custom elements.
 * The Blazor instance is passed to afterStarted as an argument
 * @param  {} blazor
 */
export async function afterStarted(blazor) {
    console.log("afterStarted: BlazorAddIn.lib.module.js():");
    console.log("afterStarted: Entering function");
}