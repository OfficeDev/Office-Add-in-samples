/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using System.Diagnostics;

using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

using Blazor.Word.AddIn.Client.Model;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.Word.AddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class Home : ComponentBase, IAsyncDisposable
{
    private HostInformation hostInformation = new();

    [Inject]
    private IJSRuntime JSRuntime { get; set; } = default!;

    private IJSObjectReference JSModule { get; set; } = default!;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            hostInformation = await JSRuntime.InvokeAsync<HostInformation>("Office.onReady");

            Debug.WriteLine("Hit OnAfterRenderAsync in Home.razor.cs!");
            Console.WriteLine("Hit OnAfterRenderAsync in Home.razor.cs in Console!");
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Home.razor.js");

            if (hostInformation.IsInitialized)
            {
                StateHasChanged();
            }
        }
    }

    [JSImport("insertParagraph", "Home")]
    internal static partial Task InsertParagraph();

    // Static JSImport method to call into Home.razor.js (synchronous - returns string, not Promise)
    [JSImport("sayHelloFromJs", "Home")]
    internal static partial string SayHelloFromJsFunction(string name);

    [JSInvokable]
    public static async Task<string> SayHelloHome(string name)
    {
        Console.WriteLine($"Invoking static SayHelloHome {name}");

        // Import the Home module (collocated JS file)
        await JSHost.ImportAsync("Home", "../Pages/Home.razor.js");

        // Call the JS function and get the result (synchronous call)
        var result = SayHelloFromJsFunction(name);

        var renderMode = OperatingSystem.IsBrowser() ? "InteractiveWebAssembly" : "Unknown";

        return $"{result} from the {renderMode} Home Page!";
    }

    public async ValueTask DisposeAsync()
    {
        if (JSModule is not null)
        {
            try
            {
                await JSModule.DisposeAsync();
            }
            catch (JSDisconnectedException)
            {
                // Circuit is already gone; JS-side resources are already cleaned up.
            }
        }
    }
}