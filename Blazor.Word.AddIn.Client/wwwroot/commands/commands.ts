/**
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
console.log("Loading command.js");

// ---------------------------------------------------------------------------
// Core helpers
// ---------------------------------------------------------------------------

/**
 * Wraps a ribbon command action with the required try/catch/finally shell
 * and ensures event.completed() is always called.
 */
async function withOfficeEvent(
  event: Office.AddinCommands.Event,
  name: string,
  action: () => Promise<void>
): Promise<void> {
  try {
    console.log(`Running ${name}`);
    await action();
  } catch (error) {
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
 * Invokes a [JSInvokable] method that returns Task<string> and appends
 * the result to the Word document body.
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
  await insertText(result);
}

/**
 * Invokes a [JSInvokable] method that returns Task (void).
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
 * Appends text to the end of the Word document body.
 *
 * @param text - The text to insert into the document
 */
async function insertText(text: string): Promise<void> {
  await Word.run(async (context) => {
    context.document.body.insertText(text, Word.InsertLocation.end);
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Ribbon command handlers
// ---------------------------------------------------------------------------

async function insertTextInWord(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "insertTextInWord", async () => {
    await Word.run(async (context) => {
      context.document.body.insertText("Hello World from TypeScript", Word.InsertLocation.end);
      await context.sync();
    });
  });
}

async function callBlazorWasm(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "callBlazorWasm", () =>
    invokeDotNetString("wasm", "SayHelloWASM", "Blazor Fan")
  );
}

async function callBlazorServer(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "callBlazorServer", () =>
    invokeDotNetString("server", "SayHelloServer", "Blazor Fan")
  );
}

async function callBlazorOnHome(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "callBlazorOnHome", () =>
    invokeDotNetString("wasm", "SayHelloHome", "Blazor Fan")
  );
}

async function callBlazorOnCounter(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "callBlazorOnCounter", () =>
    invokeDotNetString("server", "SayHelloCounter", "Blazor Fan")
  );
}

async function callBlazorPrepareDocument(event: Office.AddinCommands.Event): Promise<void> {
  await withOfficeEvent(event, "callBlazorPrepareDocument", () =>
    invokeDotNetVoid("wasm", "PrepareDocument")
  );
}

// ---------------------------------------------------------------------------
// Associate handlers with manifest XML function names
// ---------------------------------------------------------------------------
Office.actions.associate("insertTextInWord", insertTextInWord);
Office.actions.associate("callBlazorWasm", callBlazorWasm);
Office.actions.associate("callBlazorServer", callBlazorServer);
Office.actions.associate("callBlazorOnHome", callBlazorOnHome);
Office.actions.associate("callBlazorOnCounter", callBlazorOnCounter);
Office.actions.associate("callBlazorPrepareDocument", callBlazorPrepareDocument);
