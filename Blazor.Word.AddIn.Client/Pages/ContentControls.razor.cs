/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

using Blazor.Word.AddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.Word.AddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class ContentControls : ComponentBase, IAsyncDisposable
{
    private bool HostInformation;
    private static bool _isImported = false;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            try
            {
                await JSHost.ImportAsync("ContentControls", "../Pages/ContentControls.razor.js");
                _isImported = true;
                Console.WriteLine($"Imported ContentControls module");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing ContentControls module: {ex.Message}");
            }

            HostInformation = await OfficeUtilities.IsRunningInHostAsync();
            Console.WriteLine($"Home HostInformation: {HostInformation}");

            if (HostInformation)
            {
                StateHasChanged();
            }
        }
    }

    // Use static JSImport methods directly
    internal static Task Setup() => SetupDocument();
    internal static Task InsertContentControls() => InsertContentControlsFunction();
    internal static Task TagContentControls() => TagContentControlsFunction();
    internal static Task ModifyContentControls() => ModifyContentControlsFunction();
    internal static Task RegisterContentControlsEvents() => RegisterContentControlsFunction();
    internal static Task DeregisterContentControlEvents() => DeregisterContentControlsFunction();
    internal static Task DeleteContentControl() => DeleteContentControlFunction();

    [JSImport("setupDocument", "ContentControls")]
    internal static partial Task SetupDocument();

    [JSImport("insertContentControls", "ContentControls")]
    internal static partial Task InsertContentControlsFunction();

    [JSImport("tagContentControls", "ContentControls")]
    internal static partial Task TagContentControlsFunction();

    [JSImport("modifyContentControls", "ContentControls")]
    internal static partial Task ModifyContentControlsFunction();

    [JSImport("registerEvents", "ContentControls")]
    internal static partial Task RegisterContentControlsFunction();

    [JSImport("deregisterEvents", "ContentControls")]
    internal static partial Task DeregisterContentControlsFunction();

    [JSImport("deleteContentControl", "ContentControls")]
    internal static partial Task DeleteContentControlFunction();

    /// <summary>
    /// Prepares the Word document by importing the required JavaScript module and setting up content controls.
    /// This method can be invoked from JavaScript via JSInterop.
    /// </summary>
    /// <returns>A task that represents the asynchronous operation.</returns>
    [JSInvokable]
    public static async Task PrepareDocument()
    {
        if (!_isImported)
        {
            await JSHost.ImportAsync("ContentControls", "../Pages/ContentControls.razor.js");
            _isImported = true;
        }

        await SetupDocument();
        await InsertContentControlsFunction();
        await TagContentControlsFunction();
        await ModifyContentControlsFunction();
    }

    public async ValueTask DisposeAsync()
    {
        if (_isImported)
        {
            try
            {
                await DeregisterContentControlsFunction();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error deregistering content controls: {ex.Message}");
            }
        }
    }
}