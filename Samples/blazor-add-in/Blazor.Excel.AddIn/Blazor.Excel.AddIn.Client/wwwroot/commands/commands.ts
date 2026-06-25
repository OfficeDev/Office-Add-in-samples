/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 *
 */

/// <reference path="./blazor.d.ts" />

console.log("Loading command.js");

// ---------------------------------------------------------------------------
// Core helpers
// ---------------------------------------------------------------------------

/**
 * Wraps a ribbon command action with the required try/catch/finally shell
 * and ensures event.completed() is always called.
 *
 * @param event - The Office add-in command event
 * @param name - The name of the command, used for logging
 * @param action - The async action to execute
 */
async function withOfficeEvent(
  event: Office.AddinCommands.Event,
  name: string,
  action: () => Promise<void>
): Promise<void> {
  try {
    console.log(`Running ${name}`);
    await action();
  } catch (error: unknown) {
    console.error(`Error in ${name}:`, error instanceof Error ? error.message : String(error));
  } finally {
    console.log(`Finish ${name}`);
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }
}

/**
 * Waits for the named bridge to be ready and returns its DotNetObjectReference.
 * Throws if the bridge times out or is not registered.
 *
 * @param bridgeName - The bridge to wait for ("wasm" or "server")
 * @param timeoutMs - Maximum time to wait for the bridge to be ready (default: 10000ms)
 */
async function resolveDotNetRef(bridgeName: string, timeoutMs = 10_000) {
  console.log(`Waiting for ${bridgeName} bridge`);

  const bridgePromise = window.dotNetReady?.[bridgeName];
  if (!bridgePromise) {
    throw new Error(`dotNetReady.${bridgeName} not found – Blazor module may not be loaded`);
  }

  await Promise.race([
    bridgePromise,
    new Promise<never>((_, reject) =>
      AbortSignal.timeout(timeoutMs).addEventListener("abort", () =>
        reject(new Error(`Timeout waiting for ${bridgeName} bridge`))
      )
    ),
  ]);

  const ref = window.dotNetRefs.get(bridgeName);
  if (!ref) {
    throw new Error(`Bridge '${bridgeName}' not found in dotNetRefs after ready signal`);
  }

  console.log(`${bridgeName} bridge resolved`);
  return ref;
}

/**
 * Invokes a [JSInvokable] instance method that returns Task<string> and inserts
 * the result into the selected Excel cell.
 *
 * @param bridgeName - The bridge name ("wasm" or "server") to look up in dotNetRefs.
 * @param methodName - The name of the [JSInvokable] method to invoke on the handler.
 * @param args - Additional arguments to pass to the .NET method.
 */
async function invokeDotNetString(bridgeName: string, methodName: string, ...args: unknown[]): Promise<void> {
  const t0 = performance.now();
  const ref = await resolveDotNetRef(bridgeName);
  const result = await ref.invokeMethodAsync<string>(methodName, ...args);
  await insertText(result);
  console.log(`${methodName} took ${(performance.now() - t0).toFixed(1)}ms`);
}

/**
 * Invokes a [JSInvokable] instance method that returns Task (void).
 *
 * @param bridgeName - The bridge name ("wasm" or "server") to look up in dotNetRefs.
 * @param methodName - The name of the [JSInvokable] method to invoke on the handler.
 * @param args - Additional arguments to pass to the .NET method.
 */
async function invokeDotNetVoid(bridgeName: string, methodName: string, ...args: unknown[]): Promise<void> {
  const t0 = performance.now();
  const ref = await resolveDotNetRef(bridgeName);
  await ref.invokeMethodAsync(methodName, ...args);
  console.log(`${methodName} took ${(performance.now() - t0).toFixed(1)}ms`);
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

// ---------------------------------------------------------------------------
// Ribbon command handlers
// ---------------------------------------------------------------------------

/**
 * Inserts the name returned by the Home Blazor page into the selected cell and highlights it.
 * @param event - The Office add-in command event
 */
async function insertNameFromHome(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "insertNameFromHome", async () => {
    await invokeDotNetString("wasm", "SayHelloHome", "Blazor Fan");
    await setColor("LightBlue");
  });
}

/**
 * Inserts the name returned by the Counter Blazor page into the selected cell and highlights it.
 * @param event - The Office add-in command event
 */
async function insertNameFromCounter(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "insertNameFromCounter", async () => {
    await invokeDotNetString("server", "SayHelloCounter", "Blazor Fan");
    await setColor("LightBlue");
  });
}

/**
 * Writes the ribbon button's event source ID into the selected cell as an ExecuteFunction proof-of-concept.
 * @param event - The Office add-in command event
 */
async function writeButtonId(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "writeButtonId", async () => {
    const message: string = "ExecuteFunction works. Button ID=" + event.source.id;
    await Excel.run(async (context: Excel.RequestContext): Promise<void> => {
      const range: Excel.Range = context.workbook.getSelectedRange();
      range.values = [[message]];
      range.getEntireColumn().format.autofitColumns();
      await context.sync();
    });
  });
}

/**
 * Calls the JSInvokable method CreateBubbles on the WASM bridge to create a bubble chart
 * after adding the data in the active worksheet.
 * @param event - The Office add-in command event
 */
async function createBubbles(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "createBubbles", () =>
    invokeDotNetVoid("wasm", "CreateBubbles")
  );
}

// ---------------------------------------------------------------------------
// Associate handlers with manifest XML function names
// ---------------------------------------------------------------------------
Office.actions.associate("insertNameFromHome", insertNameFromHome);
Office.actions.associate("insertNameFromCounter", insertNameFromCounter);
Office.actions.associate("writeButtonId", writeButtonId);
Office.actions.associate("createBubbles", createBubbles);

