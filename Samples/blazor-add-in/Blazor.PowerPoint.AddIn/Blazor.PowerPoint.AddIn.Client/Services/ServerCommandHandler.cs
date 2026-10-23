using Microsoft.JSInterop;

namespace Blazor.PowerPoint.AddIn.Client.Services;

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
}
