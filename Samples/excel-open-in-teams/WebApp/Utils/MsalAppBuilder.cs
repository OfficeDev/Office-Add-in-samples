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

using Microsoft.Extensions.DependencyInjection;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.TokenCacheProviders;
using Microsoft.Identity.Web.TokenCacheProviders.InMemory;
using System;
using System.Security.Claims;
using System.Threading.Tasks;

namespace WebApp.Utils
{
    public static class MsalAppBuilder
    {
        public static string GetAccountId(this ClaimsPrincipal claimsPrincipal)
        {
            string oid = claimsPrincipal.GetObjectId();
            string tid = claimsPrincipal.GetTenantId();
            return $"{oid}.{tid}";
        }

        public static async Task<IConfidentialClientApplication> BuildConfidentialClientApplication()
        {
            IConfidentialClientApplication clientapp = ConfidentialClientApplicationBuilder.Create(AuthenticationConfig.ClientId)
                  .WithClientSecret(AuthenticationConfig.ClientSecret)
                  .WithRedirectUri(AuthenticationConfig.RedirectUri)
                  .WithAuthority(new Uri(AuthenticationConfig.Authority))
                  .Build();

            // After the ConfidentialClientApplication is created, we overwrite its default UserTokenCache serialization with our implementation
            IMsalTokenCacheProvider memoryTokenCacheProvider = CreateTokenCacheSerializer();
            await memoryTokenCacheProvider.InitializeAsync(clientapp.UserTokenCache);
            return clientapp;
        }

        public static async Task RemoveAccount()
        {
            IConfidentialClientApplication clientapp = ConfidentialClientApplicationBuilder.Create(AuthenticationConfig.ClientId)
                  .WithClientSecret(AuthenticationConfig.ClientSecret)
                  .WithRedirectUri(AuthenticationConfig.RedirectUri)
                  .WithAuthority(new Uri(AuthenticationConfig.Authority))
                  .Build();

            // We only clear the user's tokens.
            IMsalTokenCacheProvider memoryTokenCacheProvider = CreateTokenCacheSerializer();
            await memoryTokenCacheProvider.InitializeAsync(clientapp.UserTokenCache);
            var userAccount = await clientapp.GetAccountAsync(ClaimsPrincipal.Current.GetAccountId());
            if (userAccount != null)
            {
                await clientapp.RemoveAsync(userAccount);
            }
        }


        private static IServiceProvider serviceProvider;

        private static IMsalTokenCacheProvider CreateTokenCacheSerializer()
        {
            if (serviceProvider == null)
            {
                // In memory token cache. Other forms of serialization are possible.
                // See https://github.com/AzureAD/microsoft-identity-web/wiki/asp-net 
                IServiceCollection services = new ServiceCollection();
                services.AddInMemoryTokenCaches();

                serviceProvider = services.BuildServiceProvider();
            }
            IMsalTokenCacheProvider msalTokenCacheProvider = serviceProvider.GetRequiredService<IMsalTokenCacheProvider>();
            return msalTokenCacheProvider;
        }

    }
}