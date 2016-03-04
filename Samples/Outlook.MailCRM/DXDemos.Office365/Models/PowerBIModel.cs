using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using DXDemos.Office365.Utils;

namespace DXDemos.Office365.Models
{
    public class PowerBIModel
    {
        /// <summary>
        /// Gets a resource specific access token for Power BI ("https://analysis.windows.net/powerbi/api")
        /// </summary>
        /// <returns>Access Token string</returns>
        private static async Task<string> getAccessToken()
        {
            // fetch from stuff user claims
            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst(SettingsHelper.ClaimTypeObjectIdentifier).Value;

            // setup app info for AuthenticationContext
            var clientCredential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.ClientSecret);
            var userIdentifier = new UserIdentifier(userObjectId, UserIdentifierType.UniqueId);

            // create auth context (note: no token cache leveraged)
            AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.AzureADAuthority);

            // get access token for Power BI
            var token = await authContext.AcquireTokenAsync(SettingsHelper.PowerBIResourceId, clientCredential, new UserAssertion(userObjectId, UserIdentifierType.UniqueId.ToString()));
            return token.AccessToken; 
        }

        /// <summary>
        /// Gets all datasets for the user
        /// </summary>
        /// <returns>List of PowerBIDataset</returns>
        public static async Task<List<PowerBIDataset>> GetDatasets()
        {
            List<PowerBIDataset> datasets = new List<PowerBIDataset>();
            var token = await getAccessToken();
            var baseAddress = new Uri("https://api.powerbi.com/beta/myorg/");
            using (var client = new HttpClient{ BaseAddress = baseAddress })
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                client.DefaultRequestHeaders.Add("Accept", "application/json; odata=verbose");
                using (var response = await client.GetAsync("datasets"))
                {
                    string responseString = await response.Content.ReadAsStringAsync();
                    JObject oResponse = JObject.Parse(responseString);
                    datasets = oResponse.SelectToken("datasets").ToObject<List<PowerBIDataset>>();
                }
            }

            return datasets;
        }

        /// <summary>
        /// Gets a specific dataset based on id
        /// </summary>
        /// <param name="id">Guid id of dataset</param>
        /// <returns>PowerBIDataset</returns>
        public static async Task<PowerBIDataset> GetDataset(Guid id)
        {
            PowerBIDataset dataset = null;
            var token = await getAccessToken();
            var baseAddress = new Uri("https://api.powerbi.com/beta/myorg/");
            using (var client = new HttpClient { BaseAddress = baseAddress })
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                client.DefaultRequestHeaders.Add("Accept", "application/json; odata=verbose");
                using (var response = await client.GetAsync(String.Format("datasets/{0}", id.ToString())))
                {
                    string responseString = await response.Content.ReadAsStringAsync();
                    JObject oResponse = JObject.Parse(responseString);
                }
            }

            return dataset;
        }

        /// <summary>
        /// Creates a dataset, including tables/columns
        /// </summary>
        /// <param name="dataset">PowerBIDataset</param>
        /// <returns>Guid id of the new dataset</returns>
        public static async Task<Guid> CreateDataset(PowerBIDataset dataset)
        {
            var token = await getAccessToken();
            var baseAddress = new Uri("https://api.powerbi.com/beta/myorg/");
            using (var client = new HttpClient{ BaseAddress = baseAddress })
            {
                var content = new StringContent(JsonConvert.SerializeObject(dataset).Replace("\"id\":\"00000000-0000-0000-0000-000000000000\",", ""), System.Text.Encoding.Default, "application/json");
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                using (var response = await client.PostAsync("datasets", content))
                {
                    string responseString = await response.Content.ReadAsStringAsync();
                    JObject oResponse = JObject.Parse(responseString);
                    dataset.id = new Guid(oResponse.SelectToken("id").ToString());
                }
            }

            return dataset.id;
        }

        /// <summary>
        /// !!!!!!!!!!!! THIS IS CURRENTLY UNSUPPORTED !!!!!!!!!!!!
        /// Deletes a dataset
        /// </summary>
        /// <param name="dataset">Guid id of the dataset</param>
        /// <returns>bool indicating success</returns>
        public static async Task<bool> DeleteDataset(Guid dataset)
        {
            bool success = false;
            var token = await getAccessToken();
            var baseAddress = new Uri("https://api.powerbi.com/beta/myorg/");
            using (var client = new HttpClient { BaseAddress = baseAddress })
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                using (var response = await client.DeleteAsync(String.Format("datasets/{0}", dataset.ToString())))
                {
                    string responseString = await response.Content.ReadAsStringAsync();
                    success = true;
                }
            }

            return success;
        }

        /// <summary>
        /// Clear all data our of a given table of a dataset
        /// </summary>
        /// <param name="dataset">Guid dataset igd</param>
        /// <param name="table">string table name</param>
        /// <returns>bool indicating success</returns>
        public static async Task<bool> ClearTable(Guid dataset, string table)
        {
            bool success = false;
            var token = await getAccessToken();
            var baseAddress = new Uri("https://api.powerbi.com/beta/myorg/");
            using (var client = new HttpClient { BaseAddress = baseAddress })
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                using (var response = await client.DeleteAsync(String.Format("datasets/{0}/tables/{1}/rows", dataset.ToString(), table)))
                {
                    string responseString = await response.Content.ReadAsStringAsync();
                    success = true;
                }
            }

            return success;
        }

        /// <summary>
        /// Adds rows to a given table and dataset in Power BI
        /// </summary>
        /// <param name="dataset">PowerBIDataset</param>
        /// <param name="table">PowerBITable</param>
        /// <param name="rows">List<Dictionary<string, object>></param>
        /// <returns></returns>
        public static async Task<bool> AddTableRows(Guid dataset, string table, List<Dictionary<string, object>> rows)
        {
            bool success = false;
            var token = await getAccessToken();
            var baseAddress = new Uri("https://api.powerbi.com/beta/myorg/");
            using (var client = new HttpClient { BaseAddress = baseAddress })
            {
                //build the json post by looping through the rows and columns for each row
                string json = "{\"rows\": [";
                foreach (var row in rows)
                {
                    //process each column on the row
                    json += "{";
                    foreach (var key in row.Keys)
                    {
                        json += "\"" + key + "\": \"" + row[key].ToString() + "\",";
                    }
                    json = json.Substring(0, json.Length - 1) + "},";
                }
                json = json.Substring(0, json.Length - 1) + "]}";
                var content = new StringContent(json, System.Text.Encoding.Default, "application/json");
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                client.DefaultRequestHeaders.Add("Accept", "application/json");
                using (var response = await client.PostAsync(String.Format("datasets/{0}/tables/{1}/rows", dataset.ToString(), table), content))
                {
                    string responseString = await response.Content.ReadAsStringAsync();
                    success = true;
                }
            }

            return success;
        }
    }

    public class PowerBIDataset
    {
        public Guid id { get; set; }
        public string name { get; set; }
        public List<PowerBITable> tables { get; set; }
    }

    public class PowerBIColumn
    {
        public string name { get; set; }
        public string dataType { get; set; }
    }

    public class PowerBITable
    {
        public string name { get; set; }
        public List<PowerBIColumn> columns { get; set; }
    }

    public class PowerBITableRef
    {
        public Guid datasetId { get; set; }
        public string tableName { get; set; }
    }

    public class PowerBITableRows : PowerBITableRef
    {
        public List<Dictionary<string, object>> rows;
    }
}
