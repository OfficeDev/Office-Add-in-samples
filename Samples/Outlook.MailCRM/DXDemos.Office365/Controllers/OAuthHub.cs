using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.AspNet.SignalR;
using DXDemos.Office365.Models;

namespace DXDemos.Office365.Controllers
{
    public class OAuthHub : Hub
    {
        public void Initialize()
        {

        }
        public void OAuthComplete(string clientID, UserModel user)
        {
            Clients.Client(clientID).oAuthComplete(user);
        }
    }
}