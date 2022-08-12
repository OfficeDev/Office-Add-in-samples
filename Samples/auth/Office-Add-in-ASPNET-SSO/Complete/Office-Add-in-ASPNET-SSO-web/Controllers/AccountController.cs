using Microsoft.Owin.Security.OpenIdConnect;
//using Microsoft.Owin.Security;
//using Microsoft.Owin.Host.SystemWeb;
using OfficeAddinSSOWeb;
//using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication;
using Newtonsoft.Json;
using Microsoft.Identity.Web;
using Microsoft.Graph;
using Microsoft.Identity.Client;
//using System.Web;

namespace OfficeAddinSSOWeb.Controllers
{
    public class AccountController : Controller
    {

        [Route("Account/SignIn")]
        [AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
       public void SignIn()
        {
            var code = HttpContext.Session.GetString("OpenIdConnect");
            var redirectUrl = Url.Action(nameof(AuthorizeComplete));
            ViewBag.redirectUrl = redirectUrl;
        }

        [Route("Account/AuthorizeComplete")]

        public async Task<ActionResult> AuthorizeComplete()
        {
            var code = HttpContext.Session.GetString("OpenIdConnect");
            var spaAuthCode = HttpContext.Session.GetString("Spa_Auth_Code");

            ViewBag.SpaAuthCode = spaAuthCode;

            return View("AuthorizeComplete");
        }
    }
}
