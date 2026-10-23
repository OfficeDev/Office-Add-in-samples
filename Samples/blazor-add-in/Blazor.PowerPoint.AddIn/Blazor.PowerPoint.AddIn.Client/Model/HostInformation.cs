namespace Blazor.PowerPoint.AddIn.Client.Model
{
    public record HostInformation
    {
        public string? Host { get; init; }
        public string? Platform { get; init; }
        public bool IsInitialized => Host != null && Platform != null;
    }
}
