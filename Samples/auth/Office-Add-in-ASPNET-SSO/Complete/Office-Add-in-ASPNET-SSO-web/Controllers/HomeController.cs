// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using OfficeAddinSSOWeb.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Graph.ExternalConnectors;

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
        public ActionResult RootAuth()
        {           
            return View("AuthorizeComplete");
        }

        [Route("Home")]
        [Route("Home/Index")]
        [AllowAnonymous]
        public ActionResult Index()
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