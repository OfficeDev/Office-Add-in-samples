using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace DXDemos.Office365.IdToken.Models
{
    /// <summary>
    /// Everything in this folder is an exact reference from https://msdn.microsoft.com/en-us/library/office/fp179819.aspx on how to validate an Exchange identity token
    /// </summary>
    public static class JsonTokenDecoder
    {
        public static DecodedJsonToken Decode(IdentityTokenRequest rawToken)
        {
            string[] tokenParts = rawToken.token.Split('.');

            if (tokenParts.Length != 3)
            {
                throw new ApplicationException("Token must have three parts separated by '.' characters.");
            }

            string encodedHeader = tokenParts[0];
            string encodedPayload = tokenParts[1];
            string signature = tokenParts[2];

            string decodedHeader = Base64UrlEncoder.Decode(encodedHeader);
            string decodedPayload = Base64UrlEncoder.Decode(encodedPayload);

            JavaScriptSerializer serializer = new JavaScriptSerializer();

            Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
            Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

            return new DecodedJsonToken(header, payload, signature);
        }
    }
}