// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
using Office_Add_in_ASPNET_SSO_WebAPI.Models;
using System;
using System.IdentityModel.Tokens;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Controllers
{
	public class AzureADAuthController : Controller
	{
		// The URL that auth should redirect to after a successful login.
		Uri loginRedirectUri => new Uri(Url.Action(nameof(Authorize), "AzureADAuth", null, Request.Url.Scheme));

		/// <summary>
		/// Logs the user into Microsoft 365.
		/// </summary>
		/// <returns>A redirect to the Microsoft 365 login page.</returns>
		public async Task<ActionResult> Login()
		{
			if (string.IsNullOrEmpty(Settings.AzureADClientId) || string.IsNullOrEmpty(Settings.AzureADClientSecret))
			{
				ViewBag.Message = "Please set your client ID and client secret in the Web.config file";
				return View();
			}

			ConfidentialClientApplicationBuilder clientBuilder = ConfidentialClientApplicationBuilder
				.Create(Settings.AzureADClientId)
				.WithClientSecret(Settings.AzureADClientSecret);
				                                                         ;
			ConfidentialClientApplication clientApp = (ConfidentialClientApplication)clientBuilder.Build();

			// Generate the parameterized URL for Azure login.
			string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
			var urlBuilder = clientApp.GetAuthorizationRequestUrl(graphScopes);
			urlBuilder.WithRedirectUri(loginRedirectUri.ToString());
			urlBuilder.WithAuthority(Settings.AzureADAuthority);

			var authUrl = await urlBuilder.ExecuteAsync(System.Threading.CancellationToken.None);

			// Redirect the browser to the login page, then come back to the Authorize method below.
			return Redirect(authUrl.ToString());
		}

		/// <summary>
		/// Gets IdToken from implicit flow and sends it to main add-in window.
		/// </summary>
		/// <returns>The default view.</returns>
		public async Task<ActionResult> Authorize()
		{

			ConfidentialClientApplicationBuilder clientBuilder = ConfidentialClientApplicationBuilder.Create(Settings.AzureADClientId);
			clientBuilder.WithClientSecret(Settings.AzureADClientSecret);
			clientBuilder.WithRedirectUri(loginRedirectUri.ToString());
			clientBuilder.WithAuthority(Settings.AzureADAuthority);

			ConfidentialClientApplication clientApp = (ConfidentialClientApplication)clientBuilder.Build();
			string[] graphScopes = { "Files.Read.All" };

			// Get and save the token.
			var authResultBuilder = clientApp.AcquireTokenByAuthorizationCode(
				graphScopes,
				Request.Params["code"]   // The auth 'code' parameter from the Azure redirect.
			);

            try
            {
                var authResult = await authResultBuilder.ExecuteAsync();
                ViewBag.AccessToken = authResult.AccessToken;
            }
            catch (Exception e)
            {
                ViewBag.Error = e.Message;
            }
			
            return View();
		}
	}
}
