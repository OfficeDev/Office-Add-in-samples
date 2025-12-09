/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Blazor.Excel.AddIn.Client.Model;

using Microsoft.AspNetCore.Components;

using Microsoft.JSInterop;

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace Blazor.Excel.AddIn.Client.Pages;

public partial class Weather : ComponentBase
{
    private HostInformation hostInformation = new();

    [Inject, AllowNull]
    private IJSRuntime JSRuntime { get; set; }
    private IJSObjectReference JSModule { get; set; } = default!;
    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            hostInformation = await JSRuntime.InvokeAsync<HostInformation>("Office.onReady");

            Debug.WriteLine("Hit OnAfterRenderAsync in Weather.razor.cs!");
            Console.WriteLine("Hit OnAfterRenderAsync in Weather.razor.cs in Console!");
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Weather.razor.js");

            if (hostInformation.IsInitialized)
            {
                StateHasChanged();
            }
        }
    }

    private WeatherForecast[]? forecasts;

    public bool IsLoading
    {
        get
        {
            return forecasts is null;
        }
    }


    protected override async Task OnInitializedAsync()
    {
        await GetWeatherData();
    }

    private async Task RefreshButton() =>
        await GetWeatherData();

    /// <summary>
    /// Function to create a new slide in the Excel presentation.
    /// </summary>
    private async Task CreateSlideButton() =>
        await JSModule.InvokeVoidAsync("createWeatherSlide");

    private async Task GetWeatherData()
    {
        forecasts = null;

        // Simulate asynchronous loading to demonstrate streaming rendering
        await Task.Delay(500);

        var startDate = DateOnly.FromDateTime(DateTime.Now);
        var summaries = new[] { "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching" };
        forecasts = [.. Enumerable.Range(1, 5).Select(index => new WeatherForecast
        {
            Date = startDate.AddDays(index),
            TemperatureC = Random.Shared.Next(-20, 55),
            Summary = summaries[Random.Shared.Next(summaries.Length)]
        })];
    }

    private async Task CopyButton()
    {
        if (forecasts is null) return;
        Console.WriteLine(forecasts);
        IEnumerable<object[]> res = [.. forecasts.Select(x => new object[] { x.Date, x.TemperatureC, x.TemperatureF, x.Summary ?? "None" })];
        await JSModule.InvokeVoidAsync("copyButton", res);
    }
}
