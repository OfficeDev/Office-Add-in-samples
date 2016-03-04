using DXDemos.Office365.Models;
using DXDemos.Office365.Utils;
using Microsoft.AspNet.SignalR;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace DXDemos.Office365.Controllers
{
    public class OAuthController : Controller
    {
        [Route("OAuth/AuthCode/{userid}/{signalrRef}/")]
        public async Task<ActionResult> AuthCode(string userid, string signalrRef)
        {
            //Request should have a code from AAD and an id that represents the user in the data store
            if (Request["code"] == null)
                return RedirectToAction("Error", "Home", new { error = "Authorization code not passed from the authentication flow" });
            else if (String.IsNullOrEmpty(userid))
                return RedirectToAction("Error", "Home", new { error = "User reference code not passed from the authentication flow" });

            //get access token using the authorization code
            var token = await TokenHelper.GetAccessTokenWithCode(userid.ToLower(), signalrRef, Request["code"], SettingsHelper.O365UnifiedAPIResourceId);

            //get the user from the datastore
            var idString = userid.ToLower();
            var user = DocumentDBRepository<UserModel>.GetItem("Users", i => i.id == idString);
            if (user == null)
                return RedirectToAction("Error", "Home", new { error = "User placeholder does not exist" });

            //get the user's details (name, email)
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token.access_token);
            client.DefaultRequestHeaders.Add("Accept", "application/json");
            using (HttpResponseMessage response = await client.GetAsync(new Uri("https://graph.microsoft.com/beta/me", UriKind.Absolute)))
            {
                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    JObject oResponse = JObject.Parse(json);
                    user.display_name = oResponse.SelectToken("displayName").ToString();
                    user.email_address = oResponse.SelectToken("mail").ToString();
                    user.initials = UserController.getInititials(user.display_name, user.email_address);
                }
                else
                    return RedirectToAction("Error", "Home", new { error = "AAD Graph service call failed" });
            }   

            //get the user's profile pic
            var outlookToken = await TokenHelper.GetAccessTokenWithRefreshToken(token.refresh_token, SettingsHelper.OutlookResourceId);
            client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + outlookToken.access_token);
            client.DefaultRequestHeaders.Add("Accept", "application/json");
            var url = "https://outlook.office365.com/api/beta/me/userphoto/$value";
            using (HttpResponseMessage response = await client.GetAsync(new Uri(url, UriKind.Absolute)))
            {
                if (response.IsSuccessStatusCode)
                {
                    var stream = await response.Content.ReadAsStreamAsync();
                    byte[] bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, (int)stream.Length);
                    user.picture = "data:image/jpeg;base64, " + Convert.ToBase64String(bytes);
                }
            }   

            //update the user with the refresh token and other details we just acquired
            user.refresh_token = token.refresh_token;
            await DocumentDBRepository<UserModel>.UpdateItemAsync("Users", idString, user);

            //notify the client through the hub
            var hubContext = GlobalHost.ConnectionManager.GetHubContext<OAuthHub>();
            hubContext.Clients.Client(signalrRef).oAuthComplete(user);

            //return view successfully
            return View();
        }
    }
}