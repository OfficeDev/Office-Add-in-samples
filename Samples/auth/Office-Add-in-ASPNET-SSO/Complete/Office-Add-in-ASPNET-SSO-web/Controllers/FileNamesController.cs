// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;

namespace OfficeAddinSSOWeb.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [RequiredScope("access_as_user")]
    public class FileNamesController : Controller
    {
        public FileNamesController(ITokenAcquisition tokenAcquisition, GraphServiceClient graphServiceClient, IOptions<MicrosoftGraphOptions> graphOptions)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphServiceClient = graphServiceClient;
            _graphOptions = graphOptions;
        }
        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly GraphServiceClient _graphServiceClient;
        private readonly IOptions<MicrosoftGraphOptions> _graphOptions;

        [HttpGet]
        public async Task<IActionResult> Get()
        {
            try
            {
                // Get list of first 10 file names from user's OneDrive root folder.
                var filelist = await _graphServiceClient.Me.Drive.Root.Children.Request().Select(u => new
                {
                    u.Name
                }).Top(10).GetAsync();

               // Map result to just the file names.
               List<string> files = new List<string>();
                foreach (var file in filelist)
                {
                    files.Add(file.Name);
                }

                return Ok(files);
            }
            catch (MsalException ex)
            {
                return StatusCode((int)HttpStatusCode.Unauthorized, "An authentication error occurred while acquiring a token for the Microsoft Graph API.\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException challengeException)
                {
                    _tokenAcquisition.ReplyForbiddenWithWwwAuthenticateHeader(_graphOptions.Value.Scopes.Split(' '),
                        challengeException.MsalUiRequiredException);
                }
                else
                {
                    return StatusCode((int)HttpStatusCode.BadRequest, "An error occurred while calling the Microsoft Graph API.\n" + ex.Message);
                }
            }
            
            return StatusCode((int)HttpStatusCode.InternalServerError);
        }
    }
}
