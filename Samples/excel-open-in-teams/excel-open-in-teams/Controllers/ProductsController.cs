using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using excel_open_in_teams.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Identity.Web;
using excel_open_in_teams.Helpers;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Net.Http;

namespace excel_open_in_teams.Controllers
{
    [Authorize]
    public class ProductsController : Controller
    {
        private readonly ITokenAcquisition _tokenAcquisition;

        // Keeps track of an in memory db for dev and testing purposes
        private readonly IProductData db;

        public ProductsController(ITokenAcquisition tokenAcquisition)
        {
            //Create in memory product data
            this.db = new InMemoryProductData();
            this._tokenAcquisition = tokenAcquisition;
        }

        /// <summary>
        /// Gets the products page which will display the product sales data
        /// </summary>
        /// <returns></returns>
        [Authorize]
        public ActionResult Products()
        {
            // Construct model with product data and return it with the view
            var model = db.GetAll();
            return View(model);
        }

        /// <summary>
        /// Step 1: After user chooses to "Open in Teams" this action is called. Calls the Graph API to
        /// get a list of Teams to return so that the user can select which team they want to open.
        /// </summary>
        /// <returns></returns>
        [AuthorizeForScopes(Scopes = new[] { "Team.ReadBasic.All" })]
        public async Task<ActionResult> TeamsList()
        {
            try
            {
                string[] scopes = new[] { "Team.ReadBasic.All" };
                // Following call will use Team.ReadBasic.All scope
                string jsonResponse = await GraphAPIHelper.CallGraphAPIGet(_tokenAcquisition, scopes, "https://graph.microsoft.com/v1.0/me/joinedTeams");
                ViewBag.TeamsReady = false;
                TeamQueryResponse json = JsonConvert.DeserializeObject<TeamQueryResponse>(jsonResponse);

                // construct Team list data to send back in view
                List<SelectListItem> items = new List<SelectListItem>();
                foreach (var entry in json.Teams)
                {
                    items.Add(new SelectListItem { Text = entry.Name, Value = entry.Id, Selected = false });
                }

                ViewBag.teamIDSelection = items;
                return View();
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message + ": More details: " + ex.InnerException;
                return View("Error");
            }
        }

        /// <summary>
        /// Step 2: Users selected a team (passed in TeamList), so call Graph API to get list of channels for that team.
        /// Return channels to the view so user can select which channel.
        /// </summary>
        /// <param name="TeamList">Contains the team that was selected out of the origina list</param>
        /// <returns></returns>
        [AuthorizeForScopes(Scopes = new[] { "Channel.ReadBasic.All" })]
        public async Task<ActionResult> ChannelsListForTeam(string teamIDSelection)
        {
            try
            {
                //Get channels for given team ID and return them
                // Following call will use Channel.ReadBasic.All scope
                string url = $"https://graph.microsoft.com/v1.0/teams/" + teamIDSelection + "/channels";
                string jsonResponse = await GraphAPIHelper.CallGraphAPIGet(_tokenAcquisition, new[] { "Channel.ReadBasic.All"}, url);
                Channels json = JsonConvert.DeserializeObject<Channels>(jsonResponse);

                System.Collections.Generic.List<SelectListItem> items = new System.Collections.Generic.List<SelectListItem>();
                foreach (var entry in json.Value)
                {
                    //Embed required channel data that we need later for each channel stored in the list
                    Channel c = new Channel() { Id = entry.Id, Name = entry.Name ,TeamId = teamIDSelection};
                    var listValue = JsonConvert.SerializeObject(c);
                    items.Add(new SelectListItem { Text = entry.Name, Value = listValue, Selected = false });
                }

                ViewBag.channelSelection = items;
                return View();
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message + ": More details: " + ex.InnerException;
                return View("Error");
            }
        }

        /// <summary>
        /// Step 3: Upload the spreadsheet now that we have the team and channel to send to.
        /// Upload the spreadsheet to the OneDrive for the selected team.
        /// Then create a message in the channel with a link to the spreadsheet to open.
        /// </summary>
        /// <param name="ChannelList">Contains the channel to create the message in.</param>
        /// <returns></returns>
        [AuthorizeForScopes(Scopes = new[] { "Team.ReadBasic.All", "Files.ReadWrite.All" })]
        public async Task<ActionResult> UploadSpreadsheet(string channelSelection)
        {
            try
            {
                // Deserialize the JSON channel information from the selection
                Channel channel = JsonConvert.DeserializeObject<Channel>(channelSelection);
                string fileName = "productdata.xlsx";

                // Build the spreadsheet
                SpreadsheetBuilder s = new SpreadsheetBuilder();
                var spreadsheetBytes = s.CreateSpreadsheet("ProductSales", db.GetAll());

                // Upload spreadsheet to the Team channel's OneDrive
                FileCreated file = await UploadSpreadsheetToOneDrive(channel.TeamId, channel.Id, channel.Name, fileName, spreadsheetBytes);

                // Create a new message in channel linking to the new spreadsheet file
                Message msg = await CreateChannelMessage(channel.TeamId, channel.Id, file, fileName);

                ViewBag.redirect = msg.webUrl; //pass along the Message redirect url to the new view.
                return View("UploadToTeams");
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message + ": More details: " + ex.InnerException;
                return View("Error");
            }
        }


        private async Task<FileCreated> UploadSpreadsheetToOneDrive(string teamID, string channelID, string channelName, string fileName, byte[] spreadsheetBytes)
        {
            try
            {
                // Construct url to get name of OneDrive file folder from team and channel
                string url = "https://graph.microsoft.com/v1.0/teams/" + teamID + "/channels/" + channelID + "/filesFolder";

                // Following call will use Team.ReadBasic.All scope
                string jsonResponse = await GraphAPIHelper.CallGraphAPIGet(_tokenAcquisition,new[] {"Team.ReadBasic.All" }, url);
                ChannelFolder json = JsonConvert.DeserializeObject<ChannelFolder>(jsonResponse);

                // Construct url to upload file to OneDrive on Graph
                url = "https://graph.microsoft.com/v1.0/drives/" + json.ParentReference.DriveID + "/items/root:/" + channelName + "/" + fileName + ":/content";

                // Following call will use Files.ReadWrite.All scope
                jsonResponse = await GraphAPIHelper.CallGraphAPIWithBody(_tokenAcquisition,new[] {"Files.ReadWrite.All" }, url, HttpMethod.Put, spreadsheetBytes);

                // Deserialize and return new file metadata
                FileCreated file = JsonConvert.DeserializeObject<FileCreated>(jsonResponse);
                return file;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private async Task<Message> CreateChannelMessage(string teamID, string channelID, FileCreated file, string fileName)
        {
            // Construct url to create the message on the channel
            string url = "https://graph.microsoft.com/v1.0/teams/" + teamID + "/channels/" + channelID + "/messages";

            // Extract the id portion of the eTag
            var tagStartLoc = file.eTag.IndexOf('{');
            string eTag = file.eTag.Substring(tagStartLoc + 1);
            eTag = eTag.Substring(0, eTag.IndexOf('}'));

            // Reset file.webUrl to just the portion needed
            var startLoc = file.webUrl.IndexOf(fileName);
            file.webUrl = file.webUrl.Substring(0, startLoc + fileName.Length);

            // Construct body of message and attach link to spreadsheet
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

            // Following call will use ChannelMessage.Send scope
            // Create message with file attachment in Teams channel
            string jsonResponse = await GraphAPIHelper.CallGraphAPIWithBody(_tokenAcquisition, new[] { "ChannelMessage.Send"}, url, HttpMethod.Post, body);

            // Return metadata describing the new message
            Message msg = JsonConvert.DeserializeObject<Message>(jsonResponse);
            return msg;
        }


    }
}
