/************************************************************************************************
The MIT License (MIT)

Copyright (c) 2015 Microsoft Corporation

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
***********************************************************************************************/

using Microsoft.Identity.Client;
using Microsoft.Owin;
using Owin;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Claims;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace WebApp.Utils
{
    /// <summary>
    /// A simple custom middleware, which takes care of intercepting messages containing authorization codes, validating them, redeeming the code
    /// and saving the resulting tokens in a MSAL cache, and finally redirecting to the URL that originated the request.
    /// </summary>
    /// <seealso cref="Microsoft.Owin.OwinMiddleware" />
    public sealed class OAuth2CodeRedeemerMiddleware : OwinMiddleware
    {
        private readonly OAuth2CodeRedeemerOptions options;

        public OAuth2CodeRedeemerMiddleware(OwinMiddleware next, OAuth2CodeRedeemerOptions options)
            : base(next)
        {
            this.options = options ?? throw new ArgumentNullException("options");

            this.options = options;
        }

        public async override Task Invoke(IOwinContext context)
        {
            string code = context.Request.Query["code"];
            if (code != null)
            {
                //extract state
                string state = HttpUtility.UrlDecode(context.Request.Query["state"]);

                string signedInUserID = context.Authentication.User.FindFirst(System.IdentityModel.Claims.ClaimTypes.NameIdentifier).Value;
                HttpContextBase hcb = context.Environment["System.Web.HttpContextBase"] as HttpContextBase;

                //validate state
                CodeRedemptionData crd = OAuth2RequestManager.ValidateState(state, hcb);

                if (crd != null)
                {
                    //if valid, redeem code

                    IConfidentialClientApplication cc = await MsalAppBuilder.BuildConfidentialClientApplication();
                    AuthenticationResult result = await cc.AcquireTokenByAuthorizationCode(crd.Scopes, code).ExecuteAsync().ConfigureAwait(false);

                    //redirect to original requestor
                    context.Response.StatusCode = 302;
                    context.Response.Headers.Set("Location", crd.RequestOriginatorUrl);
                }
                else
                {
                    context.Response.StatusCode = 302;
                    context.Response.Headers.Set("Location", "/Error?message=" + "code_redeem_failed");
                }
            }
            else
            {
                await Next.Invoke(context);
            }
        }
    }

    public sealed class OAuth2CodeRedeemerOptions
    {
        public string ClientId { get; set; }
        public string RedirectUri { get; set; }
        public string ClientSecret { get; set; }
    }

    internal static class OAuth2CodeRedeemerHandler
    {
        public static IAppBuilder UseOAuth2CodeRedeemer(this IAppBuilder app, OAuth2CodeRedeemerOptions options)
        {
            app.Use<OAuth2CodeRedeemerMiddleware>(options);
            return app;
        }
    }

    #region Utils

    public class OAuth2RequestManager
    {
        // 9188040d-6c67-4c5b-b112-36a304b66dad is the GUID that indicates that the user is a consumer user from a Microsoft account. All personal account will have this tenant id.
        public const string ConsumerTenantId = "9188040d-6c67-4c5b-b112-36a304b66dad";

        private static readonly ReaderWriterLockSlim SessionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);

        /// Generate a state value using a random Guid value, the origin of the request and the scopes being requested.
        /// The state value will be consumed by the OAuth controller for validation, for specifying the corresc scopes during code redemption,
        /// and redirection after code redemption.
        /// Here we store the random Guid in the session for validation by the OAuth controller.
        private static string GenerateState(string requestUrl, HttpContextBase httpcontext, UrlHelper url, string[] scopes)
        {
            try
            {
                string stateGuid = Guid.NewGuid().ToString();
                SaveUserStateValue(stateGuid, httpcontext);

                List<String> stateList = new List<string>
                {
                    stateGuid,
                    requestUrl
                };

                // turn the scopes array into a comma separated list string
                string scopeslist = scopes[0];
                if (scopes.Count() > 1)
                {
                    for (int i = 1; i < scopes.Count(); i++)
                    {
                        scopeslist += "," + scopes[i];
                    }
                }

                stateList.Add(scopeslist);

                var formatter = new BinaryFormatter();
                var stream = new MemoryStream();
                formatter.Serialize(stream, stateList);
                var stateBits = stream.ToArray();

                return url.Encode(Convert.ToBase64String(stateBits));
            }
            catch
            {
                return null;
            }
        }

        // save the state in the session for the current user
        private static void SaveUserStateValue(string stateGuid, HttpContextBase httpcontext)
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(System.IdentityModel.Claims.ClaimTypes.NameIdentifier).Value;
            SessionLock.EnterWriteLock();
            httpcontext.Session[signedInUserID + "_state"] = stateGuid;
            SessionLock.ExitWriteLock();
        }

        private static string ReadUserStateValue(HttpContextBase httpcontext)
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(System.IdentityModel.Claims.ClaimTypes.NameIdentifier).Value;
            string stateGuid = string.Empty;
            SessionLock.EnterReadLock();
            stateGuid = (string)httpcontext.Session[signedInUserID + "_state"];
            SessionLock.ExitReadLock();
            return stateGuid;
        }

        // Verify that the state identifier in the state string corresponds to the GUID saved in the session for the current user
        // If the check succeeds, return the scopes to request when redeeming the code and the URL to which the app will be redirected after redemption
        public static CodeRedemptionData ValidateState(string state, HttpContextBase httpcontext)
        {
            try
            {
                var stateBits = Convert.FromBase64String(HttpUtility.UrlDecode(state));
                var formatter = new BinaryFormatter();
                var stream = new MemoryStream(stateBits);
                List<String> stateList = (List<String>)formatter.Deserialize(stream);
                var stateGuid = stateList[0];
                //TODO - cleaning up should not be necessary, I have just one entry per user
                // but at least I should do it for making the state single use
                if (stateGuid == ReadUserStateValue(httpcontext))
                {
                    string returnURL = stateList[1];
                    string[] scopes = stateList[2].Split(',');
                    return new CodeRedemptionData()
                    {
                        RequestOriginatorUrl = returnURL,
                        Scopes = scopes
                    };
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }

        public static async Task<string> GenerateAuthorizationRequestUrl(string[] scopes, IConfidentialClientApplication cca, HttpContextBase httpcontext, UrlHelper url)
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(System.IdentityModel.Claims.ClaimTypes.NameIdentifier).Value;
            string preferredUsername = ClaimsPrincipal.Current.FindFirst("preferred_username").Value;
            Uri oauthCodeProcessingPath = new Uri(httpcontext.Request.Url.GetLeftPart(UriPartial.Authority).ToString());
            string state = GenerateState(httpcontext.Request.Url.ToString(), httpcontext, url, scopes);
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            string domain_hint = (tenantID == ConsumerTenantId) ? "consumers" : "organizations";

            Uri authzMessageUri = await cca
                .GetAuthorizationRequestUrl(scopes)
                .WithRedirectUri(oauthCodeProcessingPath.ToString())
                .WithLoginHint(preferredUsername)
                .WithExtraQueryParameters(state == null ? null : "&state=" + state + "&domain_hint=" + domain_hint)
                .WithAuthority(cca.Authority)
                .ExecuteAsync(CancellationToken.None)
                .ConfigureAwait(false);

            return authzMessageUri.ToString();
        }
    }

    public class CodeRedemptionData
    {
        public string RequestOriginatorUrl { get; set; }
        public string[] Scopes { get; set; }
    }

    #endregion Utils
}