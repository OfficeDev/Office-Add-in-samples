// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Configuration;

namespace AttachmentDemoWeb.Helpers
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