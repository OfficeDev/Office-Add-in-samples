using Blazor.Excel.AddIn.Client.Pages;

using Microsoft.JSInterop;

namespace Blazor.Excel.AddIn.Client.Services;

/// <summary>
/// Command handler service for Server-side ribbon commands.
/// Methods are exposed to JavaScript via DotNetObjectReference through ServerBridge.
/// Executes on the server via SignalR.
/// </summary>
public class ServerCommandHandler
{
    [JSInvokable]
    public Task<string> SayHelloServer(string name)
    {
        Console.WriteLine($"Invoking SayHelloServer {name}");
        return Task.FromResult($"Hello Server, {name} from the InteractiveServer ServerCommandHandler!");
    }

    /// <summary>
    /// Delegates to the static Counter.SayHelloCounter method.
    /// This allows JS to call via DotNetObjectReference (targeting the Server runtime)
    /// while still going through Counter.razor.cs.
    /// </summary>
    [JSInvokable]
    public Task<string> SayHelloCounter(string name)
    {
        Console.WriteLine($"ServerCommandHandler.SayHelloCounter delegating to Counter.SayHelloCounter");
        return Counter.SayHelloCounter(name);
    }
}
