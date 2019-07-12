// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using OutlookAddinMicrosoftGraphASPNET.Helpers;
using OutlookAddinMicrosoftGraphASPNET.Models;
using System.Web.Mvc;

namespace OutlookAddinMicrosoftGraphASPNET.Controllers
{
    public class HomeController : Controller
    {
        /// <summary>
        /// Presents the user with a home page or the data retrieval page, depending on whether the user
        /// is signed in.
        /// </summary>
        /// <returns>The default view.</returns>
        public ActionResult Index()
        {
            var userAuthStateId = Settings.GetUserAuthStateId(ControllerContext.HttpContext);
            if (Data.GetUserSessionToken(userAuthStateId, Settings.AzureADAuthority) != null)
            {
                // When the user is signed in, go directly to the list of workbooks.
                return RedirectToAction("Index", "Files");
            }

            // If the user isn't signed in, go to the home page with its Connect button.
            ViewBag.StateKey = userAuthStateId;
            var token = new SessionToken();
            return View(token);
        }
    }
}
