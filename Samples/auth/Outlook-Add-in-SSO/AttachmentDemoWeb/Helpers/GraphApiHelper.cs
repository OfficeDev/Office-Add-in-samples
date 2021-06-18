// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* 
    This file provides access to OneDrive through Microsoft Graph. 
*/

using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using AttachmentDemoWeb.Models;
using Microsoft.Graph;
using System.IO;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace AttachmentDemoWeb.Helpers
{
    /// <summary>
    /// Provides methods and strings for Microsoft Graph-specific endpoints.
    /// </summary>
    internal static class GraphApiHelper
    {
        internal static async Task<HttpResponseMessage> WriteAttachmentsToOneDrive(string accessToken, SaveAttachmentRequest request)
        {
            try
            {
                GraphServiceClient graphClient = new GraphServiceClient(
                  new DelegateAuthenticationProvider(
                      (requestMessage) =>
                      {
                          // Add the OneDrive access token to each outgoing request
                          requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                          return Task.FromResult(0);
                      }));
                foreach (string attachmentId in request.attachmentIds)
                {
                    var attachment = await graphClient.Me.Messages[request.messageId].Attachments[attachmentId].Request().GetAsync();

                    // Is this a file or an Outlook item?
                    if (string.Compare(attachment.ODataType, "#microsoft.graph.itemAttachment") == 0)
                    {
                        // Re-request the attachment with the item expanded
                        var itemAttachment = await graphClient.Me.Messages[request.messageId].Attachments[attachmentId].Request()
                            .Expand("microsoft.graph.itemAttachment/item").GetAsync() as ItemAttachment;

                        // Serialize the item to JSON and save to OneDrive
                        string jsonItem = JsonConvert.SerializeObject(itemAttachment.Item);
                        MemoryStream fileStream = new MemoryStream();
                        StreamWriter sw = new StreamWriter(fileStream);
                        sw.Write(jsonItem);
                        sw.Flush();
                        fileStream.Position = 0;
                        bool success = await SaveFileToOneDrive(graphClient, itemAttachment.Name + ".json", fileStream);
                        if (!success)
                        {
                            return HttpErrorHelper.SendErrorToClient(HttpStatusCode.BadRequest, new Exception(string.Format("Could not save {0} to OneDrive", itemAttachment.Name)), null);
                        }
                    }
                    else
                    {
                        var fileAttachment = attachment as FileAttachment;

                        // For files, we can build a stream directly from ContentBytes
                        if (fileAttachment.Size < (4 * 1024 * 1024))
                        {
                            MemoryStream fileStream = new MemoryStream(fileAttachment.ContentBytes);
                            bool success = await SaveFileToOneDrive(graphClient, fileAttachment.Name, fileStream);
                            if (!success)
                            {
                                return HttpErrorHelper.SendErrorToClient(HttpStatusCode.BadRequest, new Exception(string.Format("Could not save {0} to OneDrive", fileAttachment.Name)), null);
                            }
                        }
                        else
                        {
                            // TODO: Add code here to handle larger files. See:
                            // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_createuploadsession
                            // and
                            // https://github.com/microsoftgraph/aspnet-snippets-sample/blob/master/Graph-ASPNET-46-Snippets/Microsoft%20Graph%20ASPNET%20Snippets/Models/FilesService.cs

                            return HttpErrorHelper.SendErrorToClient(HttpStatusCode.InternalServerError, new Exception("File is too large for simple upload."), null);
                        }
                    }
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
            {
                return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, e, null);
            }
            return new HttpResponseMessage(HttpStatusCode.OK);

        }

        internal static async Task<bool> SaveFileToOneDrive(GraphServiceClient client, string fileName, Stream fileContent)
        {
            string relativeFilePath = "Outlook Attachments/" + MakeFileNameValid(fileName);

            try
            {
                // This method only supports files 4MB or less
                DriveItem newItem = await client.Me.Drive.Root.ItemWithPath(relativeFilePath)
                    .Content.Request().PutAsync<DriveItem>(fileContent);
            }
            catch (ServiceException)
            {
                return false;
            }

            return true;
        }

        internal static string MakeFileNameValid(string originalFileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", originalFileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
        }
    }
}