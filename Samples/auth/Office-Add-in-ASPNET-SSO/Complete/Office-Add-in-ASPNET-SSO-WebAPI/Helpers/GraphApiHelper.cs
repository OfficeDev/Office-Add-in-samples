// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of the project.

/* 
    This file provides URLs to help get Microsoft Graph data. 
*/

using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;
using Office_Add_in_ASPNET_SSO_WebAPI.Models;
using System.Web.Http;


namespace Office_Add_in_ASPNET_SSO_WebAPI.Helpers
{
    /// <summary>
    /// Provides methods and strings for Microsoft Graph-specific endpoints.
    /// </summary>
    internal static class GraphApiHelper
    {
        // Microsoft Graph-related base URLs
        internal static string GetFilesUrl = @"https://graph.microsoft.com/v1.0/me/drive/root/children";

        // The selectedProperties parameter is a query parameter that will be added
        // to the Microsoft Graph REST API URL. If any part of it comes from user input,
        // be sure that it is sanitized so that it cannot be used in
        // a Response header injection attack.
        internal static string GetOneDriveItemNamesUrl(string selectedProperties)
        {
            // Construct URL for the names of the folders and files.
            return GetFilesUrl + selectedProperties;
        }

        internal static async Task<HttpResponseMessage> GetOneDriveFileNames(string accessToken)
        {
            // Get the names of files and folders in OneDrive by using the Microsoft Graph API.
            var fullOneDriveItemsUrl = GetOneDriveItemNamesUrl("?$select=name&$top=10");
            IEnumerable<OneDriveItem> filesResult;
            try
            {
                filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, accessToken);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
            {
                return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, e, null);
            }

            // Remove excess information from the data and send the data to the client.
            List<string> itemNames = new List<string>();
            foreach (OneDriveItem item in filesResult)
            {
                itemNames.Add(item.Name);
            }

            var requestMessage = new HttpRequestMessage();
            requestMessage.SetConfiguration(new HttpConfiguration());
            var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames);
            return response;
        }
    }
}
