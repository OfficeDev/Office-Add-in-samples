using Microsoft.AspNetCore.Mvc;
using System.Net;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;

namespace Office_Add_in_ASPNETCore_WebAPI.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [RequiredScope("access_as_user")]
    public class FilesController : Controller
    {
        public FilesController(ITokenAcquisition tokenAcquisition, GraphServiceClient graphServiceClient, IOptions<MicrosoftGraphOptions> graphOptions)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphServiceClient = graphServiceClient;
            _graphOptions = graphOptions;

        }

        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly GraphServiceClient _graphServiceClient;
        private readonly IOptions<MicrosoftGraphOptions> _graphOptions;

        // GET api/files
        [HttpGet]
        [Produces("application/json")]
        public async Task<IActionResult> Get()
        {
            List<DriveItem> result = new List<DriveItem>();
            try
            {
                var files = await _graphServiceClient.Me.Drive.Root.Children.Request()
                    .Top(10)
                    .Select(m => new { m.Name })
                    .GetAsync();

                result = files.ToList();
            }
            catch (MsalException ex)
            {
                var errorResponse = new
                {
                    message = "An authentication error occurred while acquiring a token for downstream API",
                    details = ex.Message
                };

                return StatusCode((int)HttpStatusCode.Unauthorized, Json(errorResponse));
            }
            catch (ServiceException ex)
            {
                if (ex.InnerException is MicrosoftIdentityWebChallengeUserException challengeException)
                {
                    _tokenAcquisition.ReplyForbiddenWithWwwAuthenticateHeader(_graphOptions.Value.Scopes.Split(' '),
                        challengeException.MsalUiRequiredException);
                }
                else
                {
                    var errorResponse = new
                    {
                        message = "An error occurred calling Microsoft Graph",
                        details = ex.RawResponseBody
                    };
                    return StatusCode((int)HttpStatusCode.BadRequest, Json(errorResponse));
                }
            }
            catch (Exception ex)
            {
                var errorResponse = new
                {
                    message = "An error occurred while calling the downstream API",
                    details = ex.Message
                };
                return StatusCode((int)HttpStatusCode.BadRequest, Json(errorResponse));

            }
            return Json(result);
        }
    }
}
