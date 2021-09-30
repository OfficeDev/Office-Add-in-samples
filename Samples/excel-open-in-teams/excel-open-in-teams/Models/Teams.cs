using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace excel_open_in_teams.Models
{
    public class TeamQueryResponse
    {
        [JsonProperty("@odata.context")]
        public string Context { get; set; }
        [JsonProperty("@odata.count")]
        public string Count { get; set; }
        [JsonProperty("value")]
        public TeamsList[] Teams { get; set; }

    }
    public class TeamsList
    {
        [JsonProperty("id")]
        public string Id { get; set; }
        [JsonProperty("displayName")]
        public string Name { get; set; }
        [JsonProperty("description")]
        public string Description { get; set; }
    }

    public class Channels
    {
        [JsonProperty("value")]
        public Channel[] Value { get; set; }

    }

    public class Channel
    {
        [JsonProperty("id")]
        public string Id { get; set; }
        [JsonProperty("displayName")]
        public string Name { get; set; }
        public string TeamId { get; set; }

    }

    public class ChannelFolder
    {
        [JsonProperty("id")]
        public string Id { get; set; }
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("parentReference")]
        public DriveModel ParentReference { get; set; }
    }
    public class DriveModel
    {
        [JsonProperty("driveId")]
        public string DriveID { get; set; }
    }
    public class FileCreated
    {
        public string id { get; set; }
        public string webUrl { get; set; }
        public string name { get; set; }
        public string eTag { get; set; }

    }

    public class Message
    {
        public string webUrl { get; set; }
    }
}
