using System;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web.Script.Serialization;

namespace DXDemos.Office365.IdToken.Models
{
    /// <summary>
    /// Everything in this folder is an exact reference from https://msdn.microsoft.com/en-us/library/office/fp179819.aspx on how to validate an Exchange identity token
    /// </summary>
    public static class AuthMetadata
    {
        public static X509Certificate2 GetSigningCertificate(Uri authMetadataEndpoint)
        {
            JsonAuthMetadataDocument document = GetMetadataDocument(authMetadataEndpoint);

            if (null != document.keys && document.keys.Length > 0)
            {
                JsonKey signingKey = document.keys[0];

                if (null != signingKey && null != signingKey.keyValue)
                {
                    return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
                }
            }

            throw new ApplicationException("The metadata document does not contain a signing certificate.");
        }

        public static JsonAuthMetadataDocument GetMetadataDocument(Uri authMetadataEndpoint)
        {
            ServicePointManager.ServerCertificateValidationCallback = Config.CertificateValidationCallback;

            byte[] acsMetadata;
            using (WebClient webClient = new WebClient())
            {
                acsMetadata = webClient.DownloadData(authMetadataEndpoint);
            }
            string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

            JsonAuthMetadataDocument document = new JavaScriptSerializer().Deserialize<JsonAuthMetadataDocument>(jsonResponseString);

            if (null == document)
            {
                throw new ApplicationException(String.Format("No authentication metadata document found at {0}.", authMetadataEndpoint));
            }

            return document;
        }
    }
}