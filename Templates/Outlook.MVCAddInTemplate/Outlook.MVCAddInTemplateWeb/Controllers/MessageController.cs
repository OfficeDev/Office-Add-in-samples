using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Outlook.MVCAddInTemplateWeb.Controllers
{
    public class MessageController : Controller
    {
        public ActionResult Read()
        {
            return View();
        }

        public ActionResult Edit()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }
    }
}