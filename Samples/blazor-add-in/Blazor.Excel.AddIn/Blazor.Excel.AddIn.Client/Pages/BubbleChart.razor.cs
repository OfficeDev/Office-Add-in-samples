/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

using Blazor.Excel.AddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.Excel.AddIn.Client.Pages;

/// <summary>
/// Starter class to demo how to create a bubble chart.
/// </summary>
[SupportedOSPlatform("browser")]
public partial class BubbleChart : ComponentBase
{
    private bool HostInformation;
    private static bool _isImported = false;

    private static async Task EnsureImportedAsync()
    {
        if (!_isImported)
        {
            await JSHost.ImportAsync("BubbleChart", "../Pages/BubbleChart.razor.js");
            _isImported = true;
        }
    }

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            try
            {
                await EnsureImportedAsync();
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
    internal static Task CreateTable() => CreateImportedTable();
    internal static Task CreateBubbleChart() => RunCreateChart();

    [JSImport("createTable", "BubbleChart")]
    internal static partial Task CreateImportedTable();

    [JSImport("createBubbleChart", "BubbleChart")]
    internal static partial Task RunCreateChart();

    private static readonly SemaphoreSlim _importLock = new(1, 1);

    /// <summary>
    /// Prepares the Excel workbook by importing the required JavaScript module and creating a Bubble Chart.
    /// This method can be invoked from JavaScript via JSInterop.
    /// </summary>
    /// <returns>A task that represents the asynchronous operation.</returns>
    [JSInvokable]
    public static async Task CreateBubbles()
    {
        await EnsureImportedAsync();

        await CreateImportedTable();
        await RunCreateChart();
    }
}