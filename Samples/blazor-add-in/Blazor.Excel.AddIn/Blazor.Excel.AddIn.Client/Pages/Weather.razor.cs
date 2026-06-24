/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Blazor.Excel.AddIn.Client.Model;
using Blazor.Excel.AddIn.Client.Services;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Runtime.Versioning;

namespace Blazor.Excel.AddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class Weather : ComponentBase, IAsyncDisposable
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
            Console.WriteLine("Hit OnAfterRenderAsync in Weather.razor.cs in Console!");
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Weather.razor.js");

            if (_hostInformation)
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
        if (forecasts is null || JSModule is null) return;
        IEnumerable<object[]> res = [.. forecasts.Select(x => new object[] { x.Date, x.TemperatureC, x.TemperatureF, x.Summary ?? "None" })];
        await JSModule.InvokeVoidAsync("copyButton", res);
    }

    public async ValueTask DisposeAsync()
    {
        if (JSModule is not null)
            await JSModule.DisposeAsync();
    }
}
