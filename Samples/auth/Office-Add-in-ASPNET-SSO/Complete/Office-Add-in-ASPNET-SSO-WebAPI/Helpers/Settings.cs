// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System;
using System.Configuration;
using System.Web;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Helpers
{
	/// <summary>
	/// Provides management of basic user and web application authentication and authorization information. 
	/// </summary>
	public static class Settings
	{
        public static string AzureADAuthority = ConfigurationManager.AppSettings["ida:Authority"];
        public static string AzureADClientId = ConfigurationManager.AppSettings["ida:ClientID"];
		public static string AzureADClientSecret = ConfigurationManager.AppSettings["ida:Password"];
		//public static string AzureADLogoutAuthority = @"https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=";
		public static string GraphApiResource = @"https://graph.microsoft.com/";
        
    }
}
