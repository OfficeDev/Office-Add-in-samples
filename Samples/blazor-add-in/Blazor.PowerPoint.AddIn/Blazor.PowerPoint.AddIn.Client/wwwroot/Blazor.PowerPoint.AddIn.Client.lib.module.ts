
/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

// NOTE: window.signalDotNetReady, window.dotNetReady, and window.dotNetRefs
// are defined in an inline <script> in App.razor so they are available before
// blazor.web.js loads. This ensures both the Server and WebAssembly interactive
// runtimes can call signalDotNetReady when their bridge components render.

/**
 * JavaScript Initializers
 * 
 * JavaScript (JS) initializers execute logic before and after a Blazor app loads.
 * JS initializers are useful in the following scenarios:
 * 
 * - Customizing how a Blazor app loads.
 * - Initializing libraries before Blazor starts up.
 * - Configuring Blazor settings.
 * 
 * To define a JS initializer, add a JS module to the project named {NAME}.lib.module.js, 
 * where the {NAME} placeholder is the assembly name, library name, or package identifier.
 * 
 * Place the file in the project's web root, which is typically the wwwroot folder.
 */

console.log("Loading Blazor.PowerPoint.AddIn.Client.lib.module.js");

/***
 * JavaScript initializers
 * https://learn.microsoft.com/en-us/aspnet/core/blazor/fundamentals/startup?preserve-view=true#javascript-initializers
 */

/**
 * beforeWebAssemblyStart(options, extensions): 
 * 
 * Called before the Interactive WebAssembly runtime is started.
 * Receives the Blazor options (options) and any extensions (extensions) added during publishing. For example, options can specify the use of a custom boot resource loader.
 * @param  {} options
 * @param  {} extensions
 */
export async function beforeWebAssemblyStart(options: any, extensions: any) {
  console.log("We are now entering function: beforeWebAssemblyStart");

  Office.onReady((info) => {
    // Check that we loaded into PowerPoint.
    if (info.host === Office.HostType.PowerPoint) {

      console.log("We are now hosting in PowerPoint.");

      // Set the startup behavior of the add-in so that it loads whenever the PowerPoint document is loaded.
      Office.addin.setStartupBehavior(Office.StartupBehavior.load);

      // Register an event handler for the onChanged event.
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
      console.log("A handler has been registered for the onChanged event.");
    }
    else {
      console.log("We are now hosting in The Browser (of your choice).");
    }
    console.log("Office onReady.");
  });
}

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event: any) {
  console.log("Change type of event: " + event.type);
}

/**
 * beforeWebStart(options): 
 * 
 * Called before the Blazor Web App starts.
 * For example, beforeWebStart is used to customize the loading process, logging level, and other options.
 * Receives the Blazor Web options (options).
 * @param  {} options
 */
export async function beforeWebStart(options: any) {
  console.log("We are now entering function: beforeWebStart");
}

/**
 * beforeServerStart(options, extensions): 
 * 
 * Called before the first Server runtime is started.
 * Receives SignalR circuit start options (options) and any extensions (extensions) added during publishing.
 * @param  {} options
 * @param  {} extensions
 */
export async function beforeServerStart(options: any, extensions: any) {
  console.log("We are now entering function: beforeServerStart");
}

/**
 * afterWebStarted(blazor): 
 * 
 * Called after all beforeWebStart promises resolve.
 * For example, afterWebStarted can be used to register Blazor event listeners and custom event types.
 * The Blazor instance is passed to afterWebStarted as an argument (blazor).
 * @param  {} blazor
 */
export async function afterWebStarted(blazor: any) {
  console.log("We are now entering function: afterWebStarted");
}

/**
 * afterServerStarted(blazor): 
 * 
 * Called after the first Interactive Server runtime is started.
 * @param  {} blazor
 */
export async function afterServerStarted(blazor: any) {
  console.log("We are now entering function: afterServerStarted");
}

/**
 * afterWebAssemblyStarted(blazor): 
 * 
 * Called after the Interactive WebAssembly runtime is started.
 * @param  {} blazor
 */
export async function afterWebAssemblyStarted(blazor: any) {
  console.log("We are now entering function: afterWebAssemblyStarted");

  // Note: The actual readiness signals are sent by the WasmBridge and ServerBridge
  // components after they render, each registering their DotNetObjectReference
  // via signalDotNetReady(name, dotNetRef).
  console.log("WebAssembly runtime started - waiting for bridge components to signal readiness");
}