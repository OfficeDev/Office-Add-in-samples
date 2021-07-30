using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using WebApp.Utils;
using Newtonsoft.Json;
using WebApp.Models;

namespace WebApp.Controllers
{
  
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        
        public ActionResult SpreadsheetTime()
        {
            SpreadsheetBuilder s = new SpreadsheetBuilder();
            var test = s.GenerateBasicSpreadsheet();

            return View("UploadToTeams");
        }

        [Authorize]
        
        public async Task<ActionResult> ChannelChosen(string ChannelList)
        {
            //Build a basic spreadsheet
            SpreadsheetBuilder s = new SpreadsheetBuilder();
            var spreadsheetBytes = s.GenerateBasicSpreadsheet();

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

            url = "https://graph.microsoft.com/v1.0/drives/" + json.ParentReference.DriveID + "/items/root:/" + channelName + "/"+ fileName + ":/content";

            ////upload file to root of Drive
            jsonResponse = await CallGraphAPI(scopes, url, HttpMethod.Put, spreadsheetBytes);
            FileCreated file = JsonConvert.DeserializeObject<FileCreated>(jsonResponse);

            url = "https://graph.microsoft.com/v1.0/teams/" + teamID + "/channels/"+ channelID + "/messages";

            var tagStartLoc = file.eTag.IndexOf('{');
            string eTag = file.eTag.Substring(tagStartLoc + 1);
            eTag = eTag.Substring(0, eTag.IndexOf('}'));

            var startLoc = file.webUrl.IndexOf(fileName);
            file.webUrl = file.webUrl.Substring(0, startLoc+fileName.Length);

            string body = @"{
                    ""body"": {
                        ""contentType"": ""html"",
                        ""content"": ""Here's the latest budget. <attachment id=\""";
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

        public async Task<ActionResult> TeamChosen(string TeamList)
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
            return View("PickChannel");

        }


        [Authorize]

        public async Task<ActionResult> OpenInTeams()
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

        [Authorize]
        public ActionResult Financial()
        {
            Financial[] f = new Financial[2];

            f[0] = new Financial();

            f[0].Title = "Ad expenses";
            f[0].Actual = "45,000";
            f[0].Budget = "47,000";
            f[0].Variance = "-2,000";

            f[1] = new Financial();
            f[1].Title = "Operating expenses";
            f[1].Actual = "103,000";
            f[1].Budget = "100,000";
            f[1].Variance = "3,000";

            ViewBag.Financials = f;
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
/// configures a body that is plan text (string)
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
            } catch (Exception ex)
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

        [Authorize]
        public async Task<ActionResult> OpenInTeamsAction()
        {
            string[] scopes = { "Team.ReadBasic.All" };
            //TBD for now assume we are passed this channel id: 19:c89a90cae3e64cf29dbf5e4ef1ee1f6f@thread.tacv2

            //get file folder for the channel
            string jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/teams/64c91dec-f6b8-46e1-9aa7-a5e61f4f2f5c/channels/19:c89a90cae3e64cf29dbf5e4ef1ee1f6f@thread.tacv2/filesFolder", HttpMethod.Get);

            //upload file to root of Drive

            jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/drives/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT/items/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT/root:/test.txt:/content", HttpMethod.Put);

            ///  string jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/teams/64c91dec-f6b8-46e1-9aa7-a5e61f4f2f5c/channels");
            /// https://graph.microsoft.com/v1.0/teams/64c91dec-f6b8-46e1-9aa7-a5e61f4f2f5c/channels/19:c89a90cae3e64cf29dbf5e4ef1ee1f6f@thread.tacv2/filesFolder
            /// https://graph.microsoft.com/v1.0/drives/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT

            ///"driveId": "b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT",
            ////drives/{drive-id}/items/{parent-id}:/{filename}:/content
            ///PUT https://graph.microsoft.com/v1.0/drives/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT/items/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT:/test.txt:/content
            ///
            /// or is it PUT https://graph.microsoft.com/v1.0/drives/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT/items/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT/root:/test.txt:/content
            ///jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/me/joinedTeams");

            return View();
        }

        [Authorize]
        public async Task<ActionResult> BrowseChannels(string teamID)
        {
            //Get channels for given team ID and return them
            string[] scopes = { "Team.ReadBasic.All" };
            string jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/teams/64c91dec-f6b8-46e1-9aa7-a5e61f4f2f5c/channels", HttpMethod.Get);
            Channels json = JsonConvert.DeserializeObject<Channels>(jsonResponse);

            System.Collections.Generic.List<SelectListItem> items = new System.Collections.Generic.List<SelectListItem>();
            foreach(var entry in json.Value)
            {
                items.Add(new SelectListItem { Text = entry.Name, Value = entry.Id+","+entry.Name, Selected = false });
            }

            ViewBag.TeamID = teamID;
            ViewBag.TeamsReady = false;
            ViewBag.ChannelsReady = true;
            ViewBag.ChannelsInfo = items;

            return View("OpenInTeams");


            //get channel info
            ///teams/{id}/channels/{id}/filesFolder

            //jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/teams/64c91dec-f6b8-46e1-9aa7-a5e61f4f2f5c/channels");
            //ViewBag.Channels = jsonResponse;
            //return View("OpenInTeams");
        }

        [Authorize]
        public async Task<ActionResult> UploadFile(string teamID, string channelID)
        {
            string[] channelInfo = channelID.Split(',');
            string[] scopes = { "Channel.ReadBasic.All", "Files.ReadWrite.All" };

            string jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/teams/" + teamID + "/channels/" + channelInfo[0] + "/filesFolder", HttpMethod.Get);

            
            string url = "https://graph.microsoft.com/v1.0/drives/" + jsonResponse + "/root:/"+channelInfo[1]+"/test.txt:/content";

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Put, url);

            //construct body of file
            string postData = "firstone=" + channelID;
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] byte1 = encoding.GetBytes(postData);

            // Set the content type of the data being posted.
            StringContent content = new StringContent("this is a test");

            request.Content = content;
            //Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            //request.Content.Headers.Add("Content-Length", byte1.Length.ToString());

            //System.IO.Stream newStream = Request.GetBufferedInputStream();

           // newStream.Write(byte1, 0, byte1.Length);

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
                Response.Write("error: An error has occurred acquiring the token from cache. Details: " + ex.Message);
            }
            if (result != null)
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                HttpResponseMessage response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    string jsonResult = await response.Content.ReadAsStringAsync();
                    Response.Write(jsonResult);
                }
                else
                {
                    Response.Write("error: 'An error has occurred calling the Graph API");
                }
            }
            ViewBag.TeamsReady = false;
            ViewBag.ChannelsReady = false;
            return View("OpenInTeams");
        }



        [Authorize]
        public async Task<ActionResult> OpenInTeamsOld()
        {
            /// string jsonResponse = await CallGraphAPI(scopes, "https://graph.microsoft.com/v1.0/teams/64c91dec-f6b8-46e1-9aa7-a5e61f4f2f5c/channels");
            /// https://graph.microsoft.com/v1.0/teams/64c91dec-f6b8-46e1-9aa7-a5e61f4f2f5c/channels/19:c89a90cae3e64cf29dbf5e4ef1ee1f6f@thread.tacv2/filesFolder
            /// https://graph.microsoft.com/v1.0/drives/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT

            /// "driveId": "b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT",
            /// /drives/{drive-id}/items/{parent-id}:/{filename}:/content
            /// PUT https://graph.microsoft.com/v1.0/drives/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT/items/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT:/test.txt:/content
            ///
            /// or is it PUT https://graph.microsoft.com/v1.0/drives/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT/items/b!_-qzB5l19k6zVbEisakxMnk9CJ4a2zFKipcKcjFC5eId7K59lfk-TLE1XvZgHJsT/root:/test.txt:/content

            string[] scopes = { "Team.ReadBasic.All" };
            string jsonResponse = await CallGraphAPI(scopes,"https://graph.microsoft.com/v1.0/me/joinedTeams", HttpMethod.Get);
            ViewBag.TeamsReady = false;
            TeamQueryResponse json = JsonConvert.DeserializeObject<TeamQueryResponse>(jsonResponse);

            System.Collections.Generic.List<SelectListItem> items = new System.Collections.Generic.List<SelectListItem>();
            foreach (var entry in json.Teams)
            {
                items.Add(new SelectListItem { Text = entry.Name, Value = entry.Id, Selected = true });
            }
          
            ViewBag.TeamsReady = true;
            ViewBag.ChannelsReady = false;
            ViewBag.TeamsInfo = items;

          
            return View();

          
        }

        [Authorize]
        public ActionResult About()
        {
            ViewBag.Name = ClaimsPrincipal.Current.FindFirst("name").Value;
            ViewBag.AuthorizationRequest = string.Empty;

            // The object ID claim will only be emitted for work or school accounts at this time.
            Claim oid = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier");
            ViewBag.ObjectId = oid == null ? string.Empty : oid.Value;

            // The 'preferred_username' claim can be used for showing the user's primary way of identifying themselves
            ViewBag.Username = ClaimsPrincipal.Current.FindFirst("preferred_username").Value;

            // The subject or nameidentifier claim can be used to uniquely identify the user
            ViewBag.Subject = ClaimsPrincipal.Current.FindFirst("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/nameidentifier").Value;
            return View();
        }

        [Authorize]
		[HttpGet]
        public async Task<ActionResult> SendMail()
        {
            // Before we render the send email screen, we use the incremental consent to obtain and cache the access token with the correct scopes
            IConfidentialClientApplication app = await MsalAppBuilder.BuildConfidentialClientApplication();
            var account = await app.GetAccountAsync(ClaimsPrincipal.Current.GetAccountId());
            string[] scopes = { "Mail.Send" };

            try
            {
				// try to get an already cached token
				await app.AcquireTokenSilent(scopes, account).ExecuteAsync().ConfigureAwait(false);
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilentAsync.
                // This indicates you need to call AcquireTokenAsync to acquire a token
                Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    // Build the auth code request Uri
                    string authReqUrl = await OAuth2RequestManager.GenerateAuthorizationRequestUrl(scopes, app, HttpContext, Url);
                    ViewBag.AuthorizationRequest = authReqUrl;
                    ViewBag.Relogin = "true";
                }
                catch (MsalException msalex)
                {
                    Response.Write($"Error Acquiring Token:{System.Environment.NewLine}{msalex}");
                }
            }
            catch (Exception ex)
            {
                Response.Write($"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}");
            }

            return View();
        }

        [Authorize]
        [HttpPost]
        public async Task<ActionResult> SendMail(string recipient, string subject, string body)
        {
            string messagetemplate = @"{{
  ""Message"": {{
    ""Subject"": ""{0}"",
    ""Body"": {{
                ""ContentType"": ""Text"",
      ""Content"": ""{1}""
    }},
    ""ToRecipients"": [
      {{
        ""EmailAddress"": {{
          ""Address"": ""{2}""
        }}
}}
    ],
    ""Attachments"": []
  }},
  ""SaveToSentItems"": ""false""
}}
";
            string message = string.Format(messagetemplate, subject, body, recipient);

            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, "https://graph.microsoft.com/v1.0/me/sendMail")
            {
                Content = new StringContent(message, Encoding.UTF8, "application/json")
            };

            IConfidentialClientApplication app = await MsalAppBuilder.BuildConfidentialClientApplication();
            AuthenticationResult result = null;
            var account = await app.GetAccountAsync(ClaimsPrincipal.Current.GetAccountId());
            string[] scopes = { "Mail.Send" };

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
				ViewBag.Error = "An error has occurred acquiring the token from cache. Details: " + ex.Message;
                return View();
            }

            if (result != null)
            {
                // Use the token to send email

                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                HttpResponseMessage response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    ViewBag.AuthorizationRequest = null;
                    return View("MailSent");
                }
            }


            return View();
        }

        public async Task<ActionResult> ReadMail()
        {
            IConfidentialClientApplication app = await MsalAppBuilder.BuildConfidentialClientApplication();
            AuthenticationResult result = null;
            var account = await app.GetAccountAsync(ClaimsPrincipal.Current.GetAccountId());
            string[] scopes = { "Mail.Read" };

            try
            {
                // try to get token silently
                result = await app.AcquireTokenSilent(scopes, account).ExecuteAsync().ConfigureAwait(false);
            }
            catch (MsalUiRequiredException)
            {
                ViewBag.Relogin = "true";
                return View();
            }
            catch (Exception eee)
            {
                ViewBag.Error = "An error has occurred. Details: " + eee.Message;
                return View();
            }

            if (result != null)
            {
                // Use the token to read email
                HttpClient hc = new HttpClient();
                hc.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("bearer", result.AccessToken);
                HttpResponseMessage hrm = await hc.GetAsync("https://graph.microsoft.com/v1.0/me/messages");

                string rez = await hrm.Content.ReadAsStringAsync();
                ViewBag.Message = rez;
            }

            return View();
        }

        public void RefreshSession()
        {
            HttpContext.GetOwinContext().Authentication.Challenge(
                new AuthenticationProperties { RedirectUri = "/Home/ReadMail" },
                OpenIdConnectAuthenticationDefaults.AuthenticationType);
        }
    }
}