using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAddinMicrosoftGraphASPNET.Models
{
    /// <summary>
    /// Models the authentication status of the user.
    /// </summary>
    public class AuthState
    {
        public string stateKey { get; set; }
       
        public string authStatus { get; set; }
    }
}