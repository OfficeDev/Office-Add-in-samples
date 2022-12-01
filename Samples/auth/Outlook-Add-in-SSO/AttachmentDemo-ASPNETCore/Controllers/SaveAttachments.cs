using AttachmentDemo_ASPNETCore.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using System.Net;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;

namespace AttachmentDemo_ASPNETCore.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [RequiredScope("access_as_user")]
    public class SaveAttachmentsController : Controller
    {
        public SaveAttachmentsController(ITokenAcquisition tokenAcquisition, GraphServiceClient graphServiceClient, IOptions<MicrosoftGraphOptions> graphOptions)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphServiceClient = graphServiceClient;
            _graphOptions = graphOptions;

        }

        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly GraphServiceClient _graphServiceClient;
        private readonly IOptions<MicrosoftGraphOptions> _graphOptions;

        // POST api/saveAttachments
        [HttpPost]
        [Produces("application/json")]
        public async Task<IActionResult> Post([FromBody] SaveAttachmentsRequestModel attachmentsRequest)
        {
            try
            {
                // You can throw mock errors to test your code here.
                // MockError("MSAL");
                if (attachmentsRequest == null) return BadRequest("No attachment ids were pass to /api/saveAttachments.");
                if (attachmentsRequest.attachmentIds == null) return BadRequest("No attachment ids were pass to /api/saveAttachments.");
                if (attachmentsRequest.messageId== null) return BadRequest("No message id was passed to /api/saveAttachments.");

                foreach (string attachmentId in attachmentsRequest.attachmentIds)
                {
                    var attachment = await _graphServiceClient.Me.Messages[attachmentsRequest.messageId].Attachments[attachmentId].Request().GetAsync();
                    
                    // Is this a file or an Outlook item?
                    if (string.Compare(attachment.ODataType, "#microsoft.graph.itemAttachment") == 0)
                    {
                        // Re-request the attachment with the item expanded

                        var itemAttachment = await _graphServiceClient.Me.Messages[attachmentsRequest.messageId].Attachments[attachmentId].Request()
                            .Expand("microsoft.graph.itemAttachment/item").GetAsync() as ItemAttachment;
                        if (itemAttachment == null) return BadRequest("Error retrieving attachment.");
                        if (itemAttachment.Item == null) return BadRequest("Error retrieving attachment.");

                        // Serialize the item to JSON and save to OneDrive
                        string jsonItem = Json(itemAttachment.Item).ToString();
                        MemoryStream fileStream = new MemoryStream();
                        StreamWriter sw = new StreamWriter(fileStream);
                        sw.Write(jsonItem);
                        sw.Flush();
                        fileStream.Position = 0;

                        // This method only supports files 4MB or less
                        DriveItem newItem = await _graphServiceClient.Me.Drive.Root.ItemWithPath(itemAttachment.Name + ".json")
                            .Content.Request().PutAsync<DriveItem>(fileStream);

                    }
                    else
                    {
                        var fileAttachment = attachment as FileAttachment;

                        // For files, we can build a stream directly from ContentBytes
                        if (fileAttachment.Size < (4 * 1024 * 1024))
                        {
                            MemoryStream fileStream = new MemoryStream(fileAttachment.ContentBytes);
                            // This method only supports files 4MB or less
                            DriveItem newItem = await _graphServiceClient.Me.Drive.Root.ItemWithPath(fileAttachment.Name)
                                .Content.Request().PutAsync<DriveItem>(fileStream);
                        }
                        else
                        {
                            // TODO: Add code here to handle larger files. See:
                            // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_createuploadsession
                            // and
                            // https://github.com/microsoftgraph/aspnet-snippets-sample/blob/master/Graph-ASPNET-46-Snippets/Microsoft%20Graph%20ASPNET%20Snippets/Models/FilesService.cs
                            return StatusCode((int)HttpStatusCode.BadRequest, Json(new { Error = "File is too large for simple upload." }));
                            
                        }
                    }
                }
            }
             catch (MsalException ex)
            {
                var errorResponse = new
                {
                    message = "An authentication error occurred while acquiring a token for downstream API",
                    details = ex.Message,
                    code = ex.ErrorCode
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
            return StatusCode((int)HttpStatusCode.OK);
        }

        /// <summary>
        /// Used for mocking MSAL and Microsoft Graph errors in your controllers.
        /// For a list of MSAL errors you may want to handle and mock, see https://learn.microsoft.com/azure/active-directory/develop/reference-aadsts-error-codes
        /// </summary>
        /// <param name="errorType">Specifies the type of mock error you want thrown.</param>
        private void MockError(string errorType)
        {
            if (errorType == "Graph")
            {
                // Mock Microsoft Graph service exception.
                Error e = new Error { Code = "invalidRange", Message = "The Microsoft Graph range was invalid." };
                ServiceException testError = new ServiceException(e, null);
                throw testError;
            }
            if (errorType == "GraphMSAL")
            {
                // Mock Microsoft Graph service exception with inner MSAL exception.
                Error e = new Error { Code = "error", Message = "An unexpected error occurred." };
                MsalUiRequiredException exMsal = new MsalUiRequiredException("scopes invalid", "The scopes are invalid");
                MicrosoftIdentityWebChallengeUserException ex = new MicrosoftIdentityWebChallengeUserException(exMsal, new string[] { "files.readwrite", "email.read" });
                ServiceException testError = new ServiceException(e, exMsal);
                throw testError;
            }
            if (errorType == "MSAL")
            {
                // Mock MSAL exception.
                MsalException e = new MsalException("AADSTS50074", "UserStrongAuthClientAuthNRequiredInterrupt - Strong authentication is required and the user did not pass the MFA challenge.");
                throw (e);
                
            }
        }   
    }
}
