// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
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