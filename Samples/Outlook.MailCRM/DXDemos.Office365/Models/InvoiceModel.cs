using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXDemos.Office365.Models
{
    public class InvoiceModel
    {
        [JsonProperty(PropertyName = "id")]
        public Guid Id { get; set; }

        [JsonProperty(PropertyName = "number")]
        public string InvoiceNumber { get; set; }

        [JsonProperty(PropertyName = "date")]
        public string InvoiceDate { get; set; }

        [JsonProperty(PropertyName = "amount")]
        public decimal Amount { get; set; }

        [JsonProperty(PropertyName = "status")]
        public string Status { get; set; }
    }
}
