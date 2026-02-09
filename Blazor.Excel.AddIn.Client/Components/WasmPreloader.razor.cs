/* Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.Excel.AddIn.Client.Components;

/// <summary>
/// Preloader component that ensures the WebAssembly runtime and Client assembly
/// are loaded and ready for JSInvokable calls from ribbon commands.
/// This component renders with InteractiveWebAssembly mode and signals readiness
/// to JavaScript when initialization completes.
/// </summary>
public partial class WasmPreloader : ComponentBase
{
    [Inject]
    private IJSRuntime JSRuntime { get; set; } = default!;

    private static bool _isReady;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender && !_isReady)
        {
            _isReady = true;
            Console.WriteLine("WasmPreloader: WebAssembly runtime and Client assembly are ready");

            // Signal to JavaScript that the WebAssembly runtime and assembly are ready
            await JSRuntime.InvokeVoidAsync("window.signalDotNetReady");
        }
    }

    /// <summary>
    /// Dummy method that can be called to verify the assembly is loaded and scanned.
    /// </summary>
    [JSInvokable]
    public static Task<string> VerifyAssemblyLoaded()
    {
        Console.WriteLine("WasmPreloader: VerifyAssemblyLoaded called");
        return Task.FromResult("Blazor.Excel.AddIn.Client assembly is loaded and ready");
    }
}
