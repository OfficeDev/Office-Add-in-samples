using DXDemos.Office365.IdToken.Models;
using DXDemos.Office365.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace DXDemos.Office365.Utils
{
    public class TokenHelper
    {
        public async static Task<Token> GetAccessTokenWithCode(string userid, string signalrRef, string code, string resource)
        {
            //Retrieve access token using authorization code
            Token token = null;
            HttpClient client = new HttpClient();
            string redirect = SettingsHelper.AppBaseUrl + "OAuth/AuthCode/";
            HttpContent content = new StringContent(String.Format(@"grant_type=authorization_code&redirect_uri={0}{1}/{2}&client_id={3}&client_secret={4}&code={5}&resource={6}", redirect, userid, signalrRef, SettingsHelper.ClientId, HttpUtility.UrlEncode(SettingsHelper.ClientSecret), code, resource));
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded");
            using (HttpResponseMessage response = await client.PostAsync("https://login.microsoftonline.com/common/oauth2/token", content))
            {
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    token = JsonConvert.DeserializeObject<Token>(json);
                }
            }
            return token;
        }

        public async static Task<Token> GetAccessTokenWithRefreshToken(string refreshToken, string resource)
        {
            //Retrieve access token using refresh token
            Token token = null;
            HttpClient client = new HttpClient();
            HttpContent content = new StringContent(String.Format(@"grant_type=refresh_token&refresh_token={0}&client_id={1}&client_secret={2}&resource={3}", refreshToken, SettingsHelper.ClientId, HttpUtility.UrlEncode(SettingsHelper.ClientSecret), resource));
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded");
            using (HttpResponseMessage response = await client.PostAsync("https://login.microsoftonline.com/common/oauth2/token", content))
            {
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    token = JsonConvert.DeserializeObject<Token>(json);
                }
            }
            return token;
        }

        public static string GetTenantIdFromToken(string accessToken)
        {
            string[] tokenParts = accessToken.Split('.');
            string encodedPayload = tokenParts[1];
            string decodedPayload = Base64UrlEncoder.Decode(encodedPayload);
            JObject oResponse = JObject.Parse(decodedPayload);
            JToken tenant_id = oResponse.SelectToken("tid");
            return tenant_id.ToString();
        }
    }
}
