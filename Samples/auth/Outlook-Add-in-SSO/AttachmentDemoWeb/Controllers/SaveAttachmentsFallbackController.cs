using System.Configuration;
using Microsoft.Identity.Client;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;
using System;
using AttachmentDemoWeb.Models;
using AttachmentDemoWeb.Helpers;

namespace AttachmentDemoWeb.Controllers
{
    public class SaveAttachmentsFallbackController : ApiController
    {
        // POST api/<controller>
        public async Task<HttpResponseMessage> Post([FromBody] SaveAttachmentRequest request)
        {
            string accessToken = Request.Headers.Authorization.ToString().Split(' ')[1];

            return await GraphApiHelper.WriteAttachmentsToOneDrive(accessToken, request);

        }
    }
}