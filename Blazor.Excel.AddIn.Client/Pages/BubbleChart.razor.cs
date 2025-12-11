/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Blazor.Excel.AddIn.Client.Model;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

namespace Blazor.Excel.AddIn.Client.Pages;

/// <summary>
/// Starter class to demo how to create a bubble chart.
/// </summary>
[SupportedOSPlatform("browser")]
public partial class BubbleChart : ComponentBase
{
    private HostInformation hostInformation = new();

    [Inject, AllowNull]
    private IJSRuntime JSRuntime { get; set; }
    private IJSObjectReference? JSModule { get; set; }

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            hostInformation = await JSRuntime.InvokeAsync<HostInformation>("Office.onReady");

            Debug.WriteLine("Hit OnAfterRenderAsync in BubbleChart.razor.cs!");
            Console.WriteLine("Hit OnAfterRenderAsync in BubbleChart.razor.cs in Console!");
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/BubbleChart.razor.js");

            if (hostInformation.IsInitialized)
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

    [JSInvokable]
    public static async Task CreateBubbles()
    {
        await JSHost.ImportAsync("BubbleChart", "../Pages/BubbleChart.razor.js");
        await CreateImportedTable();
        await RunCreateChart();
    }

    [JSInvokable]
    public static string SayHelloBubble(string name)
    {
        return $"Hello Bubble, {name} from BubbleChart!";
    }
}