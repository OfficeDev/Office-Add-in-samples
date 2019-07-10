// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System.Threading.Tasks;

namespace OutlookAddinMicrosoftGraphASPNET.Helpers
{
    /// <summary>
    /// Provides methods for Microsoft Graph-specific endpoints.
    /// </summary>
    internal static class GraphApiHelper
    {
        // Microsoft Graph-related base URLs
        internal static string GetFilesUrl = @"https://graph.microsoft.com/v1.0/me/drive/root/children";
        internal static string BaseMSGraphSearchUrl = @"https://graph.microsoft.com/v1.0/me/drive/root/microsoft.graph.search";
       // internal static string BaseItemsUrl = @"https://graph.microsoft.com/1/me/drive/items/";

        internal static string GetWorkbookSearchUrl(string selectedProperties)
        {
            // Construct URL to search OneDrive for Business for Excel workbooks                
            var workbooksSearchRelativeUrl = "(q = '.xlsx')";
            return BaseMSGraphSearchUrl + workbooksSearchRelativeUrl + selectedProperties;
        }
    }
}

