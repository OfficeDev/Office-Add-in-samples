using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXDemos.Office365.Models
{
    public class ContactModel
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        [JsonProperty(PropertyName = "address")]
        public string Address { get; set; }

        [JsonProperty(PropertyName = "work_phone")]
        public string WorkPhone { get; set; }  
  
        [JsonProperty(PropertyName = "cell_phone")]
        public string CellPhone { get; set; }

        [JsonProperty(PropertyName = "domain")]
        public string Domain { get; set; }

        [JsonProperty(PropertyName = "invoices")]
        public List<InvoiceModel> Invoices { get; set; }

        [JsonProperty(PropertyName = "notes")]
        public List<NoteModel> Notes { get; set; }

        [JsonProperty(PropertyName = "associates")]
        public List<ContactModel> Associates { get; set; }  
    }
}
