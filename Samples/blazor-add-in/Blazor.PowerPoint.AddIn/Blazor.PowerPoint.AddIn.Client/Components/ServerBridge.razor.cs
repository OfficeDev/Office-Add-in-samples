using Blazor.PowerPoint.AddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.PowerPoint.AddIn.Client.Components;

/// <summary>
/// Bridge component that registers the ServerCommandHandler with JavaScript
/// via DotNetObjectReference. Runs in InteractiveServer mode so that
/// JSInvokable methods on the handler execute on the server via SignalR.
/// </summary>
public partial class ServerBridge : ComponentBase, IDisposable
{
    [Inject]
    private IJSRuntime JSRuntime { get; set; } = default!;

    [Inject]
    private ServerCommandHandler Handler { get; set; } = default!;

    private DotNetObjectReference<ServerCommandHandler>? _dotNetRef;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            Console.WriteLine("ServerBridge: Registering ServerCommandHandler with JavaScript");
            _dotNetRef = DotNetObjectReference.Create(Handler);
            await JSRuntime.InvokeVoidAsync("window.signalDotNetReady", "server", _dotNetRef);
        }
    }

    public void Dispose()
    {
        _dotNetRef?.Dispose();
    }
}
