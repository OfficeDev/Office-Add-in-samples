
namespace DXDemos.Office365.IdToken.Models
{
    /// <summary>
    /// Everything in this folder is an exact reference from https://msdn.microsoft.com/en-us/library/office/fp179819.aspx on how to validate an Exchange identity token
    /// </summary>
    public class JsonAuthMetadataDocument
    {
        public string id { get; set; }
        public string version { get; set; }
        public string name { get; set; }
        public string realm { get; set; }
        public string serviceName { get; set; }
        public string issuer { get; set; }
        public string[] allowedAudiences { get; set; }
        public JsonKey[] keys;
        public JsonEndpoint[] endpoints;
    }

    public class JsonEndpoint
    {
        public string location { get; set; }
        public string protocol { get; set; }
        public string usage { get; set; }
    }

    public class JsonKey
    {
        public string usage { get; set; }
        public JsonKeyValue keyValue { get; set; }
    }

    public class JsonKeyValue
    {
        public string type { get; set; }
        public string value { get; set; }
    }
}