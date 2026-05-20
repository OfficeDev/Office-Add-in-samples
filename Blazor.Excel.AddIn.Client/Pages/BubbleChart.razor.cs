/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Blazor.Excel.AddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Diagnostics;

using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

namespace Blazor.Excel.AddIn.Client.Pages;

/// <summary>
/// Starter class to demo how to create a bubble chart.
/// </summary>
[SupportedOSPlatform("browser")]
public partial class BubbleChart : ComponentBase, IAsyncDisposable
{
    private bool _hostInformation;

    [Inject]
    private IJSRuntime JSRuntime { get; set; } = default!;
    private IJSObjectReference? JSModule { get; set; }

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            _hostInformation = await OfficeUtilities.IsRunningInHostAsync();

            Debug.WriteLine("Hit OnAfterRenderAsync in BubbleChart.razor.cs!");
            Console.WriteLine("Hit OnAfterRenderAsync in BubbleChart.razor.cs in Console!");
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/BubbleChart.razor.js");

            if (_hostInformation)
            {
                StateHasChanged();
            }
        }
    }

    /// <summary>
    /// Function to create the starter table as source for the bubble chart.
    /// </summary>
    private async Task CreateTable()
    {
        if (JSModule != null)
        {
            await JSModule.InvokeVoidAsync("createTable");
        }
        else
        {
            Console.Error.WriteLine("JSModule is null. Cannot create table.");
        }
    }
    /// <summary>
    /// Function to create the actual bubble chart.
    /// </summary>
    private async Task CreateBubbleChart()
    {
        if (JSModule != null)
        {
            await JSModule.InvokeVoidAsync("createBubbleChart");
        }
        else
        {
            Console.Error.WriteLine("JSModule is null. Cannot create bubble chart.");
        }
    }

    [JSImport("createTable", "BubbleChart")]
    internal static partial Task CreateImportedTable();

    [JSImport("createBubbleChart", "BubbleChart")]
    internal static partial Task RunCreateChart();

    private static bool _bubbleChartImported = false;
    private static readonly SemaphoreSlim _importLock = new(1, 1);

    [JSInvokable]
    public static async Task CreateBubbles()
    {
        if (!_bubbleChartImported)
        {
            await _importLock.WaitAsync();
            try
            {
                if (!_bubbleChartImported)
                {
                    await JSHost.ImportAsync("BubbleChart", "../Pages/BubbleChart.razor.js");
                    _bubbleChartImported = true;
                }
            }
            finally
            {
                _importLock.Release();
            }
        }
        await CreateImportedTable();
        await RunCreateChart();
    }

    [JSInvokable]
    public static string SayHelloBubble(string name)
    {
        return $"Hello Bubble, {name} from BubbleChart!";
    }

    public async ValueTask DisposeAsync()
    {
        if (JSModule is not null)
            await JSModule.DisposeAsync();
    }
}