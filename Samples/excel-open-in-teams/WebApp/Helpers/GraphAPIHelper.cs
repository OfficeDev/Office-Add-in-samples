// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text;
using Microsoft.Identity.Client;
using System.Security.Claims;
using System.Net.Http.Headers;


namespace WebApp.Utils
{
    static public class GraphAPIHelper
    {
        /// <summary>
        /// Call the Graph API using Get verb
        /// </summary>
        /// <param name="scopes">Required scopes for the call</param>
        /// <param name="url">The url to call</param>
        /// <returns>The JSON result from the call</returns>
       static public async Task<string> CallGraphAPIGet(string[] scopes, string url)
        {
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url);
            try
            {
                string accessToken = await GetAccessToken(scopes);
                return await CallGraphAPI(accessToken, request);
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        /// <summary>
        /// POST HTTP message to the Graph API that contains a plain text body.
        /// </summary>
        /// <param name="scopes">Required scopes for the call</param>
        /// <param name="url">url to POST to</param>
        /// <param name="verb">Either Post or Put</param>
        /// <param name="body">Plain text contenxt to put in body</param>
        /// <returns>The JSON result from the call</returns>
        static public async Task<string> CallGraphAPIWithBody(string[] scopes, string url, HttpMethod verb, string body = null)
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
                throw;
            }
        }

        /// <summary>
        /// POST HTTP message to the Graph API that contains a byte array body.
        /// </summary>
        /// <param name="scopes">Required scopes for the call</param>
        /// <param name="url">url to POST to</param>
        /// <param name="verb">Either Post or Put</param>
        /// <param name="body">byte array contents to put in the body</param>
        /// <returns>The JSON result from the call</returns>
        static public  async Task<string> CallGraphAPIWithBody(string[] scopes, string url, HttpMethod verb, byte[] body = null)
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
                throw;
            }
        }

        /// <summary>
        /// Gets an access token for the required scopes to make a Graph API call
        /// </summary>
        /// <param name="scopes">Requested scopes for the access token</param>
        /// <returns>Access token for the requested scopes</returns>
        static private async Task<string> GetAccessToken(string[] scopes)
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
                throw new Exception("Could not get access token for Graph API call",ex);
            }
            return result.AccessToken;
        }

        /// <summary>
        /// Runs the network HTTPS call to the Graph API and returns the results
        /// </summary>
        /// <param name="accessToken">The access token for the call</param>
        /// <param name="request">A prepared Https request to run and make the call</param>
        /// <returns>JSON result from the call</returns>
        static private async Task<string> CallGraphAPI(string accessToken, HttpRequestMessage request)
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
                        throw new Exception("Error calling Graph API. HTTP status code: " + response.StatusCode);
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception( "An error occurred calling the Microsoft Graph API.", ex);
            }
        }
    }
}