
using DXDemos.Office365.Models;
using DXDemos.Office365.Utils;
using System;
namespace DXDemos.Office365.IdToken.Models
{
    /// <summary>
    /// Everything in this folder is an exact reference from https://msdn.microsoft.com/en-us/library/office/fp179819.aspx on how to validate an Exchange identity token
    /// </summary>
    public class IdentityTokenResponse
    {
        public IdentityTokenResponse()
        {
            loginUrl = String.Format("https://login.microsoftonline.com/common/oauth2/authorize?client_id={0}&resource={1}&response_type=code&redirect_uri=", SettingsHelper.ClientId, SettingsHelper.O365UnifiedAPIResourceId);
            redirectUrl = SettingsHelper.AppBaseUrl + "OAuth/AuthCode/";
        }
        public bool validToken { get; set; }
        public bool validUser { get; set; }
        public string loginUrl { get; set; }
        public string redirectUrl { get; set; }
        public string errorMessage { get; set; }
        public IdentityToken token { get; set; }
        public UserModel user { get; set; }
    }
}