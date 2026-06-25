using Blazor.PowerPoint.AddIn.Client.Pages;

using Microsoft.JSInterop;

namespace Blazor.PowerPoint.AddIn.Client.Services;

/// <summary>
/// Command handler service for WebAssembly-side ribbon commands.
/// Methods are exposed to JavaScript via DotNetObjectReference through WasmBridge.
/// Executes in the browser's WebAssembly runtime.
/// </summary>
public class ClientCommandHandler
{
    [JSInvokable]
    public Task<string> SayHelloWASM(string name)
    {
        Console.WriteLine($"Invoking SayHelloWASM {name}");
        return Task.FromResult($"Hello WASM, {name} from the InteractiveWebAssembly ClientCommandHandler!");
    }

    /// <summary>
    /// Delegates to the static Home.SayHelloHome method.
    /// This allows JS to call via DotNetObjectReference (targeting the WebAssembly runtime)
    /// while still going through Home.razor.cs which uses JSImport to call Home.razor.js.
    /// </summary>
    [JSInvokable]
    public Task<string> SayHelloHome(string name)
    {
        Console.WriteLine($"ClientCommandHandler.SayHelloHome delegating to Home.SayHelloHome");
        return Home.SayHelloHome(name);
    }
}
