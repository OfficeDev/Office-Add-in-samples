using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace DXDemos.Office365.Utils
{
    public class SettingsHelper 
    {
        public static string AppBaseUrl
        {
            get { return ConfigurationManager.AppSettings["baseUrl"]; }
        }
        
        public static string ClientId {
            get { return ConfigurationManager.AppSettings["ida:ClientID"]; }
        }
        
        public static string ClientSecret {
            get { return ConfigurationManager.AppSettings["ida:Password"]; }
        }

        //resource IDs for various services...try to use unified API for most things
        public static string O365UnifiedAPIResourceId {
            get { return "https://graph.microsoft.com/"; }
        }

        public static string O365DiscoveryResourceId {
            get { return "https://api.office.com/discovery/"; }
        }

        public static string OutlookResourceId {
            get { return "https://outlook.office365.com/"; }
        }

        public static string AzureAdGraphResourceId {
            get { return "https://graph.windows.net"; }
        }
        
        public static string PowerBIResourceId {
            get { return "https://analysis.windows.net/powerbi/api"; }
        }

        public static string OneNoteResourceId {
            get { return "https://onenote.com/"; }
        }




        public static string O365DiscoveryServiceEndpoint {
            get { return "https://api.office.com/discovery/v1.0/me/"; }
        }

        public static string AzureADAuthority {
            get { return "https://login.microsoftonline.com/common/"; }
        }

        public static string ClaimTypeObjectIdentifier {
            get { return "http://schemas.microsoft.com/identity/claims/objectidentifier"; }
        }
    }
}