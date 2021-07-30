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

using System.Configuration;
using System.Globalization;

namespace WebApp.Utils
{
    public static class AuthenticationConfig
    {
        public const string IssuerClaim = "iss";
        public const string TenantIdClaimType = "http://schemas.microsoft.com/identity/claims/tenantid";
        public const string MicrosoftGraphGroupsApi = "https://graph.microsoft.com/v1.0/groups";
        public const string MicrosoftGraphUsersApi = "https://graph.microsoft.com/v1.0/users";
        public const string AdminConsentFormat = "https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&state={2}&redirect_uri={3}";
        public const string BasicSignInScopes = "openid profile offline_access";
        public const string NameClaimType = "name";

        /// <summary>
        /// The Client ID is used by the application to uniquely identify itself to Azure AD.
        /// </summary>
        public static string ClientId { get; } = ConfigurationManager.AppSettings["ida:ClientId"];

        /// <summary>
        /// The ClientSecret is a credential used to authenticate the application to Azure AD.  Azure AD supports password and certificate credentials.
        /// </summary>
        public static string ClientSecret { get; } = ConfigurationManager.AppSettings["ida:ClientSecret"];

        /// <summary>
        /// The Redirect Uri is the URL where the user will be redirected after they sign in.
        /// </summary>
        public static string RedirectUri { get; } = ConfigurationManager.AppSettings["ida:RedirectUri"];

        public static string AADInstance { get; } = ConfigurationManager.AppSettings["ida:AADInstance"];

        /// <summary>
        /// The authority
        /// </summary>
        public static string Authority = string.Format(CultureInfo.InvariantCulture, AADInstance, "common", "/v2.0");
    }
}