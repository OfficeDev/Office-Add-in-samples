namespace Blazor.PowerPoint.AddIn.Client.Model
{
    public record WeatherForecast
    {
        public DateOnly Date { get; init; }
        public int TemperatureC { get; init; }
        public string? Summary { get; init; }
        public int TemperatureF => 32 + (int)(TemperatureC * 9.0 / 5.0);
    }
}
