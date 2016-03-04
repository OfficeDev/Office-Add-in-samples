using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DXDemos.Office365.Controllers
{
    public class AgaveController : Controller
    {
        // GET: Agave
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult MailCRM()
        {
            return View();
        }
    }
}