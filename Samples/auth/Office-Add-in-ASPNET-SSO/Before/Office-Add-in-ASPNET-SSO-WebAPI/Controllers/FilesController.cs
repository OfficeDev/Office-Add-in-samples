using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Threading.Tasks;
using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
using Office_Add_in_ASPNET_SSO_WebAPI.Models;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Controllers
{
    public class FilesController : ApiController
    {

        // GET api/files
        public async Task<HttpResponseMessage> Get()
        {
            string accessToken = Request.Headers.Authorization.ToString().Split(' ')[1];

            return await GraphApiHelper.GetOneDriveFileNames(accessToken);
        }
    }
}
