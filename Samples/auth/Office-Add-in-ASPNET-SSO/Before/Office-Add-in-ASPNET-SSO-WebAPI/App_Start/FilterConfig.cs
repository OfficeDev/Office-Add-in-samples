using System.Web;
using System.Web.Mvc;

namespace Office_Add_in_ASPNET_SSO_WebAPI
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
