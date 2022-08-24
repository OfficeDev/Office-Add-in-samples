// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using OfficeAddinSSOWeb.Models;
using Microsoft.AspNetCore.Authorization;

namespace OfficeAddinSSOWeb.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
        


        [Route("")]
        [AllowAnonymous]
        // temp for handling the oidc redirect but change this later.
        public async Task<ActionResult> RootAuth()
        {
            var code = HttpContext.Session.GetString("OpenIdConnect");
            var spaAuthCode = HttpContext.Session.GetString("Spa_Auth_Code");

            ViewBag.SpaAuthCode = spaAuthCode;

            return View("AuthorizeComplete");
        }

        [Route("Home")]
        [Route("Home/Index")]
        [AllowAnonymous]
        public async Task<ActionResult> Index()
        {
            return View();
        }

        [AllowAnonymous]
        public IActionResult Privacy()
        {
            return View();
        }

        [AllowAnonymous]
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}