/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 *
 */

/// <reference path="./blazor.d.ts" />

console.log("Loading command.js");

/**
 * Writes the text from the Home Blazor Page to the Worksheet when highlightSelectionHome runs.
 * @param event - The Office add-in command event
 */
async function highlightSelectionHome(event: Office.AddinCommands.Event): Promise<void> {
  try {
    console.log("Running highlightSelectionHome");
    await callDotNetMethod("wasm", "SayHelloHome");

    await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
      const range: Excel.Range = context.workbook.getSelectedRange();
      range.format.fill.color = "LightBlue";
      await context.sync();
    });
  } catch (error: unknown) {
    console.error(error);
  } finally {
    console.log("Finish highlightSelectionHome");
    event.completed();
  }
}

/**
 * Writes the text from the Counter Blazor Page to the Worksheet when highlightSelectionCounter runs.
 * @param event - The Office add-in command event
 */
async function highlightSelectionCounter(event: Office.AddinCommands.Event): Promise<void> {
  try {
    console.log("Running highlightSelectionCounter");
    await callDotNetMethod("server", "SayHelloCounter");

    await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
      const range: Excel.Range = context.workbook.getSelectedRange();
      range.format.fill.color = "LightBlue";
      await context.sync();
    });
  } catch (error: unknown) {
    console.error(error);
  } finally {
    console.log("Finish highlightSelectionCounter");
    event.completed();
  }
}

/**
 * Writes the event source id to the document when ExecuteFunction runs.
 * @param event - The Office add-in command event
 */
async function writeValue(event: Office.AddinCommands.Event): Promise<void> {

  console.log("In writeValue");

  try {
    const message: string = "ExecuteFunction works. Button ID=" + event.source.id;

    await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
      const range: Excel.Range = context.workbook.getSelectedRange();
      range.values = [[message]];
      range.getEntireColumn().format.autofitColumns();
      await context.sync();
    });

    console.log("writeValue Succeeded");

  } catch (error: unknown) {
    const errorMessage: string = error instanceof Error ? error.message : String(error);

    await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
      const range: Excel.Range = context.workbook.getSelectedRange();
      const cellRange: Excel.Range = range.getCell(0, 0);
      cellRange.values = [[errorMessage]];
      await context.sync();
    });

    console.log("Error call : " + errorMessage);
  } finally {
    console.log("Finish writeValue");
    event.completed();
  }
}

/**
 * Calls the JSInvokable function CreateBubbles to create a bubble chart, after adding the data in the active worksheet.
 * @param event - The Office add-in command event
 */
async function createBubbles(event: Office.AddinCommands.Event): Promise<void> {

  console.log("Running createBubbles");

  try {
    const dotnetloaded = await preloadDotNet("wasm");

    if (dotnetloaded) {
      await DotNet.invokeMethodAsync(
        "Blazor.Excel.AddIn.Client",
        "CreateBubbles");
      console.log("Finished createBubbles");
    } else {
      console.error("createBubbles: WASM bridge not ready");
    }
  } catch (error: unknown) {
    console.error(error);
  } finally {
    console.log("Finish createBubbles");
    event.completed();
  }
}

/**
 * Writes the text from the Home Blazor Page to the Excel Cell
 * @param event - The Office add-in command event
 */
async function callBlazorOnHome(event: Office.AddinCommands.Event): Promise<void> {
  try {
    console.log("Running callBlazorOnHome");
    await callDotNetMethod("wasm", "SayHelloHome");
  } catch (error: unknown) {
    console.error(error);
  } finally {
    console.log("Finish callBlazorOnHome");
    event.completed();
  }
}

/**
 * Writes the text from the Counter Blazor Page to the Excel Cell
 * @param event - The Office add-in command event
 */
async function callBlazorOnCounter(event: Office.AddinCommands.Event): Promise<void> {
  try {
    console.log("Running callBlazorOnCounter");
    await callDotNetMethod("server", "SayHelloCounter");
  } catch (error: unknown) {
    console.error(error);
  } finally {
    console.log("Finish callBlazorOnCounter");
    event.completed();
  }
}

/**
 * Invokes a .NET method on the named DotNetObjectReference and inserts the result
 * into the selected Excel cell.
 *
 * @param {string} bridgeName - The bridge name ("wasm" or "server") to look up in dotNetRefs.
 * @param {string} methodName - The name of the [JSInvokable] method to invoke on the handler.
 */
async function callDotNetMethod(bridgeName: string, methodName: string): Promise<void> {

  const t0 = performance.now();
  console.log(`In callDotNetMethod: bridge=${bridgeName}, method=${methodName}`);

  try {
    let name = "Initializing";

    try {
      const dotnetloaded = await preloadDotNet(bridgeName);

      if (dotnetloaded === true) {
        const dotNetRef = window.dotNetRefs.get(bridgeName);

        if (!dotNetRef) {
          name = `Bridge '${bridgeName}' not found in dotNetRefs`;
          console.error(name);
        } else {
          name = "Dotnet Loaded";
          // Call [JSInvokable] instance method on the DotNetObjectReference
          name = await dotNetRef.invokeMethodAsync<string>(methodName, "Blazor Fan");
        }
      } else {
        name = "Init DotNet Failed, methodName: " + methodName;
      }
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      name = errorMessage;
      console.error("Error during DotNet invocation: " + name);
    }

    console.log(`callDotNetMethod: .NET call took ${(performance.now() - t0).toFixed(1)}ms`);
    await insertText(name);
    await setColor("yellow");
    console.log("Finished: " + name);
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    await insertText("Error: " + errorMessage);
    await setColor("red");
    console.error("Error in callDotNetMethod:", errorMessage);
  } finally {
    console.log("Finish callDotNetMethod");
  }
}

/**
 * Inserts text into the first cell of the currently selected Excel range.
 *
 * @param text - The text to insert into the cell
 */
async function insertText(text: string): Promise<void> {
  await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
    const range: Excel.Range = context.workbook.getSelectedRange();
    const cellRange: Excel.Range = range.getCell(0, 0);
    cellRange.values = [[text]];
    range.getEntireColumn().format.autofitColumns();
    await context.sync();
  });
}

/**
 * Sets the fill color of the currently selected Excel range.
 *
 * @param color - The fill color to apply
 */
async function setColor(color: string): Promise<void> {
  await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
    const range: Excel.Range = context.workbook.getSelectedRange();
    range.format.fill.color = color;
    await context.sync();
  });
}

/**
 * Waits for a specific .NET bridge component to be ready.
 *
 * Each bridge (wasm, server) has its own promise that resolves when
 * that bridge's component signals readiness via signalDotNetReady(name, dotNetRef).
 *
 * @param bridgeName - The bridge to wait for ("wasm" or "server")
 * @param timeoutMs - Maximum time to wait for the bridge to be ready (default: 10000ms)
 * @returns {Promise<boolean>} Returns true if the bridge is ready, false if timeout.
 */
async function preloadDotNet(bridgeName: string, timeoutMs: number = 10000): Promise<boolean> {
  console.log(`In preloadDotNet: waiting for ${bridgeName} bridge`);

  try {
    const bridgePromise = window.dotNetReady?.[bridgeName];

    if (!bridgePromise) {
      console.error(`dotNetReady.${bridgeName} promise not found - Blazor module may not be loaded`);
      return false;
    }

    let timeoutId: ReturnType<typeof setTimeout> | undefined;
    const timeoutPromise = new Promise<never>((_, reject) => {
      timeoutId = setTimeout(() => {
        reject(new Error(`Timeout waiting for ${bridgeName} bridge`));
      }, timeoutMs);
    });

    try {
      // Race between the bridge ready promise and a timeout
      await Promise.race([bridgePromise, timeoutPromise]);
    } finally {
      clearTimeout(timeoutId);
    }

    console.log(`${bridgeName} bridge is ready`);
    return true;

  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error("Error in preloadDotNet: " + errorMessage);
    return false;

  } finally {
    console.log("Finish preloadDotNet");
  }
}

// Associate the functions with their named counterparts in the manifest XML.
Office.actions.associate("highlightSelectionHome", highlightSelectionHome);
Office.actions.associate("highlightSelectionCounter", highlightSelectionCounter);
Office.actions.associate("writeValue", writeValue);
Office.actions.associate("createBubbles", createBubbles);
Office.actions.associate("callBlazorOnHome", callBlazorOnHome);
Office.actions.associate("callBlazorOnCounter", callBlazorOnCounter);

