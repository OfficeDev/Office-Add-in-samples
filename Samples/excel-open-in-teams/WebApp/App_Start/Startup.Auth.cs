using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Owin.Host.SystemWeb;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.Notifications;
using Microsoft.Owin.Security.OpenIdConnect;
using Owin;
using System;
using System.Security.Claims;
using System.Threading.Tasks;
using WebApp.Utils;

namespace WebApp
{
    public partial class Startup
    {
        public void ConfigureAuth(IAppBuilder app)
        {
            app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);

            app.UseCookieAuthentication(new CookieAuthenticationOptions());

            // Custom middleware initialization. This is activated when the code obtained from a code_grant is present in the querystring (&code=<code>).
            app.UseOAuth2CodeRedeemer(
                new OAuth2CodeRedeemerOptions
                {
                    ClientId = AuthenticationConfig.ClientId,
                    ClientSecret = AuthenticationConfig.ClientSecret,
                    RedirectUri = AuthenticationConfig.RedirectUri
                });

            app.UseOpenIdConnectAuthentication(
                new OpenIdConnectAuthenticationOptions
                {
                    // The `Authority` represents the v2.0 endpoint - https://login.microsoftonline.com/common/v2.0
                    Authority = AuthenticationConfig.Authority,
                    ClientId = AuthenticationConfig.ClientId,
                    RedirectUri = AuthenticationConfig.RedirectUri,
                    PostLogoutRedirectUri = AuthenticationConfig.RedirectUri,
                    Scope = AuthenticationConfig.BasicSignInScopes + " Mail.Read", // a basic set of permissions for user sign in & profile access "openid profile offline_access"
                    TokenValidationParameters = new TokenValidationParameters
                    {
                        ValidateIssuer = false,
                        // In a real application you would use IssuerValidator for additional checks, like making sure the user's organization has signed up for your app.
                        //     IssuerValidator = (issuer, token, tvp) =>
                        //     {
                        //        //if(MyCustomTenantValidation(issuer))
                        //        return issuer;
                        //        //else
                        //        //    throw new SecurityTokenInvalidIssuerException("Invalid issuer");
                        //    },
                        //NameClaimType = "name",
                    },
                    Notifications = new OpenIdConnectAuthenticationNotifications()
                    {
                        AuthorizationCodeReceived = OnAuthorizationCodeReceived,
                        AuthenticationFailed = OnAuthenticationFailed,
                        RedirectToIdentityProvider = OnRedirectToIdentityProvider,
                    },
                    // Handling SameSite cookie according to https://docs.microsoft.com/en-us/aspnet/samesite/owin-samesite
                    CookieManager = new SameSiteCookieManager(
                                     new SystemWebCookieManager())
                });
        }

        private  Task OnRedirectToIdentityProvider(RedirectToIdentityProviderNotification<OpenIdConnectMessage, OpenIdConnectAuthenticationOptions> arg)
        {
            arg.ProtocolMessage.SetParameter("myNewParameter", "its Value");
            return Task.CompletedTask;
        }

        private async Task OnAuthorizationCodeReceived(AuthorizationCodeReceivedNotification context)
        {
            // Upon successful sign in, get the access token & cache it using MSAL
            IConfidentialClientApplication clientApp = await MsalAppBuilder.BuildConfidentialClientApplication();
            AuthenticationResult result = await clientApp.AcquireTokenByAuthorizationCode(new[] { "Mail.Read" }, context.Code).ExecuteAsync();
        }

        private Task OnAuthenticationFailed(AuthenticationFailedNotification<OpenIdConnectMessage, OpenIdConnectAuthenticationOptions> notification)
        {
            notification.HandleResponse();
            notification.Response.Redirect("/Error?message=" + notification.Exception.Message);
            return Task.FromResult(0);
        }
    }
}