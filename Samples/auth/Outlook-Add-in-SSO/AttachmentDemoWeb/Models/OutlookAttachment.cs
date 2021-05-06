using Newtonsoft.Json;

namespace AttachmentDemoWeb.Models
{
    public class OutlookAttachment
    {
        [JsonProperty("@odata.type")]
        public string Type { get; set; }
        [JsonProperty("Id")]
        public string Id { get; set; }
        [JsonProperty("Name")]
        public string Name { get; set; }
        [JsonProperty("ContentBytes")]
        public string ContentBytes { get; set; }
        [JsonProperty("Size")]
        public double Size { get; set; }
    }
}