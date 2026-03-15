/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

using Blazor.Excel.AddIn.Client.Model;
using Blazor.Excel.AddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.Excel.AddIn.Client.Pages;

/// <summary>
/// Page displaying current time in EST and CET time zones.
/// </summary>
[SupportedOSPlatform("browser")]
public partial class WorldClock : ComponentBase, IDisposable
{
    private string EstTime = string.Empty;
    private string CestTime = string.Empty;
    private Timer? _timer;
    private bool _hostInformation;

    [Inject, AllowNull]
    private IJSRuntime JSRuntime { get; set; }

    private static readonly TimeZoneInfo EstTimeZone = TimeZoneInfo.FindSystemTimeZoneById("America/New_York");
    private static readonly TimeZoneInfo CestTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Europe/Berlin");

    protected override void OnInitialized()
    {
        UpdateTimes();
        StartTimer();
    }

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            try
            {
                await JSHost.ImportAsync("WorldClock", "../Pages/WorldClock.razor.js");
                Console.WriteLine("Imported WorldClock module");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error importing WorldClock module: {ex.Message}");
            }

            _hostInformation = await OfficeUtilities.IsRunningInHostAsync();
            Console.WriteLine($"WorldClock HostInformation: {_hostInformation}");

            if (_hostInformation)
            {
                StateHasChanged();
            }
        }
    }

    private void StartTimer()
    {
        _timer?.Dispose();
        _timer = new Timer(_ => InvokeAsync(() =>
        {
            UpdateTimes();
            StateHasChanged();
        }), null, TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(1));
    }

    private void UpdateTimes()
    {
        var utcNow = DateTimeOffset.UtcNow;

        var estDateTime = TimeZoneInfo.ConvertTime(utcNow, EstTimeZone);
        var cestDateTime = TimeZoneInfo.ConvertTime(utcNow, CestTimeZone);

        EstTime = estDateTime.ToString("dddd, MMMM dd, yyyy HH:mm:ss");
        CestTime = cestDateTime.ToString("dddd, MMMM dd, yyyy HH:mm:ss");
    }

    private void RefreshTime()
    {
        UpdateTimes();
    }

    /// <summary>
    /// JSInvokable method that can be called from JavaScript to force a refresh.
    /// This is useful when the ribbon button is clicked to show the taskpane.
    /// </summary>
    [JSImport("refreshPage", "WorldClock")]
    internal static partial Task RefreshPageFromRibbon();

    [JSInvokable]
    public static Task<string> SayHelloWorldClock(string name)
    {
        Console.WriteLine("Invoking SayHelloWorldClock");
        return Task.FromResult($"Hello, {name} from World Clock Page!");
    }

    public void Dispose()
    {
        _timer?.Dispose();
    }
}
