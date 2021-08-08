using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebApp.Models;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using WebApp.Utils;
using Newtonsoft.Json;

namespace WebApp.Controllers
{

    public class ProductsController : Controller
    {
        private readonly IProductData db;

        public ProductsController()
        {
            this.db = new InMemoryProductData();
        }

        [Authorize]
        // GET: Products
        public ActionResult Products()
        {
            var model = db.GetAll();
            return View(model);
        }

        [Authorize]

        public async Task<ActionResult> UploadSpreadsheet(string ChannelList)
        {
            //Build a basic spreadsheet
            SpreadsheetBuilder s = new SpreadsheetBuilder();
            var spreadsheetBytes = s.CreateSpreadsheet("ProductSales");

            //upload that file
            //get file folder for the channel
            string[] subs = ChannelList.Split(',');
            string channelID = subs[0];
            string teamID = subs[1];
            string channelName = subs[2];
            string fileName = "financials.xlsx";
            string[] scopes = { "Team.ReadBasic.All" };

            string url = "https://graph.microsoft.com/v1.0/teams/" + teamID + "/channels/" + channelID + "/filesFolder";

            string jsonResponse = await CallGraphAPI(scopes, url, HttpMethod.Get);

            ChannelFolder json = JsonConvert.DeserializeObject<ChannelFolder>(jsonResponse);

            url = "https://graph.microsoft.com/v1.0/drives/" + json.ParentReference.DriveID + "/items/root:/" + channelName + "/" + fileName + ":/content";

            ////upload file to root of Drive
            jsonResponse = await CallGraphAPI(scopes, url, HttpMethod.Put, spreadsheetBytes);
            FileCreated file = JsonConvert.DeserializeObject<FileCreated>(jsonResponse);

            url = "https://graph.microsoft.com/v1.0/teams/" + teamID + "/channels/" + channelID + "/messages";

            var tagStartLoc = file.eTag.IndexOf('{');
            string eTag = file.eTag.Substring(tagStartLoc + 1);
            eTag = eTag.Substring(0, eTag.IndexOf('}'));

            var startLoc = file.webUrl.IndexOf(fileName);
            file.webUrl = file.webUrl.Substring(0, startLoc + fileName.Length);

            string body = @"{
                    ""body"": {
                        ""contentType"": ""html"",
                        ""content"": ""Here's the product sales data for discussion. <attachment id=\""";
            body += eTag + @"\""></attachment>""
                    },
                    ""attachments"": [
                        {
                            ""id"": """;
            body += eTag + @""",
                            ""contentType"": ""reference"",
                            ""contentUrl"": """;
            body += file.webUrl + @""",
                            ""name"": """;
            body += file.name + @"""
                        }
                    ]
                }";

            //Create message with file attachment in Teams channel
            jsonResponse = await CallGraphAPI(scopes, url, HttpMethod.Post, body);

            Message msg = JsonConvert.DeserializeObject<Message>(jsonResponse);
            ViewBag.redirect = msg.webUrl;

            return View("UploadToTeams");

        }

        [Authorize]

        public async Task<ActionResult> ChannelsListForTeam(string TeamList)
        {
            //Get channels for given team ID and return them
            string[] scopes = { "Team.ReadBasic.All" };
            string url = $"https://graph.microsoft.com/v1.0/teams/" + TeamList + "/channels";
            string jsonResponse = await CallGraphAPI(scopes, url, HttpMethod.Get);
            Channels json = JsonConvert.DeserializeObject<Channels>(jsonResponse);

            System.Collections.Generic.List<SelectListItem> items = new System.Collections.Generic.List<SelectListItem>();
            foreach (var entry in json.Value)
            {
                items.Add(new SelectListItem { Text = entry.Name, Value = entry.Id + "," + TeamList + "," + entry.Name, Selected = false });
            }

            ViewBag.ChannelList = items;
            return View();

        }


        [Authorize]

        public async Task<ActionResult> TeamsList()
        {
            List<SelectListItem> items = new List<SelectListItem>();
            string[] scopes = { "Team.ReadBasic.All" };

            string jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/me/joinedTeams", HttpMethod.Get);
            ViewBag.TeamsReady = false;
            TeamQueryResponse json = JsonConvert.DeserializeObject<TeamQueryResponse>(jsonResponse);

            foreach (var entry in json.Teams)
            {
                items.Add(new SelectListItem { Text = entry.Name, Value = entry.Id, Selected = false });
            }

            ViewBag.TeamList = items;
            return View();
        }




        private async Task<string> CallGraphAPI(string[] scopes, string url, HttpMethod verb)
        {
            HttpRequestMessage request = new HttpRequestMessage(verb, url);
            try
            {
                string accessToken = await GetAccessToken(scopes);
                return await CallGraphAPI(accessToken, request);
            }
            catch (Exception ex)
            {
                return "{ error: 'An error occurred attempting to get the access token. Details: " + ex.Message + "'}";
            }
        }

        /// <summary>
        /// configures a body that is plain text (string)
        /// </summary>
        /// <param name="scopes"></param>
        /// <param name="url"></param>
        /// <param name="verb"></param>
        /// <param name="body"></param>
        /// <returns></returns>
        private async Task<string> CallGraphAPI(string[] scopes, string url, HttpMethod verb, string body = null)
        {
            HttpRequestMessage request = new HttpRequestMessage(verb, url);
            if (body != null)
            {
                ASCIIEncoding encoding = new ASCIIEncoding();

                StringContent content = new StringContent(body);
                content.Headers.ContentType.MediaType = "application/json";
                request.Content = content;
            }
            try
            {
                string accessToken = await GetAccessToken(scopes);
                return await CallGraphAPI(accessToken, request);
            }
            catch (Exception ex)
            {
                return "{ error: 'An error occurred attempting to get the access token. Details: " + ex.Message + "'}";
            }
        }

        private async Task<string> GetAccessToken(string[] scopes)
        {
            IConfidentialClientApplication app = await MsalAppBuilder.BuildConfidentialClientApplication();
            AuthenticationResult result = null;
            var account = await app.GetAccountAsync(ClaimsPrincipal.Current.GetAccountId());
            try
            {
                // try to get an already cached token
                result = await app.AcquireTokenSilent(scopes, account).ExecuteAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                /*
				 * When the user access this page (from the HTTP GET action result) we check if they have the scope "Mail.Send" and 
				 * we handle the additional consent step in case it is needed. Then, we acquire an access token and MSAL cache it for us.
				 * So in this HTTP POST action result, we can always expect a token to be in cache. If they are not in the cache, 
				 * it means that the user accessed this route via an unsual way.
				 */
                throw ex;
            }
            return result.AccessToken;
        }

        /// <summary>
        /// This one does the actual call over the network
        /// </summary>
        /// <param name="accessToken"></param>
        /// <param name="request"></param>
        /// <returns></returns>
        private async Task<string> CallGraphAPI(string accessToken, HttpRequestMessage request)
        {
            HttpClient client = new HttpClient();
            try
            {
                if (accessToken != null)
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    HttpResponseMessage response = await client.SendAsync(request);

                    if (response.IsSuccessStatusCode)
                    {
                        string jsonResult = await response.Content.ReadAsStringAsync();
                        return jsonResult;
                    }
                    else
                    {
                        return "{ error: 'An error has occurred calling the Microsoft Graph API'}";
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                return "{ error: 'An error has occurred calling the Microsoft Graph API. Details: " + ex.Message + "'}";
            }
        }


        private async Task<string> CallGraphAPI(string[] scopes, string url, HttpMethod verb, byte[] body = null)
        {
            HttpRequestMessage request = new HttpRequestMessage(verb, url);
            if (body != null)
            {
                ASCIIEncoding encoding = new ASCIIEncoding();
                System.Net.Http.ByteArrayContent content = new ByteArrayContent(body);
                request.Content = content;
            }

            try
            {
                string accessToken = await GetAccessToken(scopes);
                return await CallGraphAPI(accessToken, request);
            }
            catch (Exception ex)
            {
                return "{ error: 'An error occurred attempting to get the access token. Details: " + ex.Message + "'}";
            }
        }


    }
}