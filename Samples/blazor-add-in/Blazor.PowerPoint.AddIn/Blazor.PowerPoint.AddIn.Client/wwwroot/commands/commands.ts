/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
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
 * the result into the currently selected PowerPoint slide as a text box.
 *
 * @param bridgeName - The bridge name ("wasm" or "server") to look up in dotNetRefs.
 * @param methodName - The name of the [JSInvokable] method to invoke on the handler.
 * @param args - Additional arguments to pass to the .NET method.
 */
async function invokeDotNetString(bridgeName: string, methodName: string, ...args: unknown[]): Promise<void> {
  const t0 = performance.now();
  const ref = await resolveDotNetRef(bridgeName);
  const result = await ref.invokeMethodAsync<string>(methodName, ...args);
  console.log(`${methodName} took ${(performance.now() - t0).toFixed(1)}ms`);
  await insertSlideText(result);
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

// ---------------------------------------------------------------------------
// PowerPoint helpers
// ---------------------------------------------------------------------------

/**
 * Inserts text into a text box on the currently selected PowerPoint slide.
 *
 * @param text - The text to insert into the slide
 * @param options - Optional positioning options for the text box
 */
async function insertSlideText(
  text: string,
  options?: { left?: number; top?: number; height?: number; width?: number }
): Promise<void> {
  const { left = 255, top = 50, height = 50, width = 450 } = options ?? {};

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const textBox = slide.shapes.addTextBox(text, { left, top, height, width });

    textBox.fill.setSolidColor("white");
    textBox.lineFormat.color = "black";
    textBox.lineFormat.weight = 1;
    textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;

    textBox.textFrame.textRange.paragraphFormat.horizontalAlignment =
      PowerPoint.ParagraphHorizontalAlignment.center;

    await context.sync();
  });
}

/**
 * Navigates PowerPoint to the last slide in the presentation.
 */
async function goToLastSlide(): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(asyncResult.error.message));
      } else {
        resolve();
      }
    });
  });
}

/**
 * Inserts a base64-encoded image into the currently selected slide.
 *
 * @param base64Image - The base64-encoded image data to insert
 */
async function insertImage(base64Image: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      base64Image,
      { coercionType: Office.CoercionType.Image },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(asyncResult.error.message));
        } else {
          resolve();
        }
      }
    );
  });
}

/**
 * Removes all default placeholder shapes from a slide to create a blank canvas.
 *
 * @param shapes - The shape collection of the target slide
 */
async function removeSlidePlaceholders(shapes: PowerPoint.ShapeCollection): Promise<void> {
  shapes.load("items");
  await shapes.context.sync();

  for (let i = shapes.items.length - 1; i >= 0; i--) {
    shapes.items[i]?.delete();
  }
  await shapes.context.sync();
}

// ---------------------------------------------------------------------------
// Ribbon command handlers
// ---------------------------------------------------------------------------

/**
 * Inserts a "Hello World" text box into the currently selected PowerPoint slide.
 * Demonstrates basic Office JavaScript API usage without Blazor interop.
 *
 * @param event - The Office add-in command event
 */
async function insertTextInPowerPoint(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "insertTextInPowerPoint", async () => {
    await insertSlideText("Hello World", { top: 25, width: 250 });
  });
}

/**
 * Calls SayHelloWASM on the WASM bridge (ClientCommandHandler) and inserts the result into the slide.
 *
 * @param event - The Office add-in command event
 */
async function callBlazorWasm(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "callBlazorWasm", () =>
    invokeDotNetString("wasm", "SayHelloWASM", "Blazor Fan")
  );
}

/**
 * Calls SayHelloServer on the server bridge (ServerCommandHandler) and inserts the result into the slide.
 *
 * @param event - The Office add-in command event
 */
async function callBlazorServer(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "callBlazorServer", () =>
    invokeDotNetString("server", "SayHelloServer", "Blazor Fan")
  );
}

/**
 * Calls SayHelloHome on the WASM bridge and inserts the result into the slide.
 * Delegates to the Home page's JSImport collocated module via ClientCommandHandler.
 *
 * @param event - The Office add-in command event
 */
async function callBlazorHome(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "callBlazorHome", () =>
    invokeDotNetString("wasm", "SayHelloHome", "Blazor Fan")
  );
}

// ---------------------------------------------------------------------------
// Associate handlers with manifest XML function names
// ---------------------------------------------------------------------------
Office.actions.associate("insertTextInPowerPoint", insertTextInPowerPoint);
Office.actions.associate("callBlazorWasm", callBlazorWasm);
Office.actions.associate("callBlazorServer", callBlazorServer);
Office.actions.associate("callBlazorHome", callBlazorHome);
