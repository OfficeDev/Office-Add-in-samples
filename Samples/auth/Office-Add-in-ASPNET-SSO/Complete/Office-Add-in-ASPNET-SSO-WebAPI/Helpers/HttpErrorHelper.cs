using System;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Helpers
{
    internal static class HttpErrorHelper
    {
        internal static HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
        {
            HttpError error;
            if (e != null)
            {
                error = new HttpError(e, true);
            }
            else
            {
                error = new HttpError(message);
            }
            var requestMessage = new HttpRequestMessage();
            var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
            return errorMessage;
        }
    }
}