using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXDemos.Office365.Models
{
    public class UserModel
    {
        [JsonProperty(PropertyName = "id")]
        public string id { get; set; }

        [JsonProperty(PropertyName = "hash")]
        public string hash { get; set; }

        [JsonProperty(PropertyName = "refresh_token")]
        public string refresh_token { get; set; }

        [JsonProperty(PropertyName = "display_name")]
        public string display_name { get; set; }

        [JsonProperty(PropertyName = "email_address")]
        public string email_address { get; set; }

        [JsonProperty(PropertyName = "initials")]
        public string initials { get; set; }

        [JsonProperty(PropertyName = "picture")]
        public string picture { get; set; }
    }
}
