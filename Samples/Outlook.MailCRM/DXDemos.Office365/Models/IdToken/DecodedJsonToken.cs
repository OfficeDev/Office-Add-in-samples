using DXDemos.Office365.Utils;
using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace DXDemos.Office365.IdToken.Models
{
    /// <summary>
    /// Everything in this folder is an exact reference from https://msdn.microsoft.com/en-us/library/office/fp179819.aspx on how to validate an Exchange identity token
    /// </summary>
    public class DecodedJsonToken : IDisposable
    {
        private readonly Dictionary<string, string> headerClaims;
        private readonly Dictionary<string, string> payloadClaims;
        private readonly Dictionary<string, string> appContext;

        private readonly string signature;

        public DecodedJsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
        {

            // We'll start out assuming that the token is invalid.
            this.IsValid = false;

            // Set the private dictionaries that contain the claims.
            this.headerClaims = header;
            this.payloadClaims = payload;
            this.signature = signature;

            // If there is no "appctx" claim in the token, throw an ApplicationException.
            if (!this.payloadClaims.ContainsKey(AuthClaimTypes.AppContext))
            {
                throw new ApplicationException(String.Format("The {0} claim is not present.", AuthClaimTypes.AppContext));
            }

            appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload[AuthClaimTypes.AppContext]);


            // Validate the header fields.
            this.ValidateHeader();

            // Determine if the token is within its valid time.
            this.ValidateLifetime();

            // Validate that the token was sent to the correct URL.
            //this.ValidateAudience();

            // Validate the token version.
            this.ValidateVersion();

            // Make sure that the appctx contains an authentication
            // metadata location.
            this.ValidateMetadataLocation();

            // If the token passes all of the validation checks, then we
            // can assume that it is valid.
            this.IsValid = true;
        }

        public string Audience
        {
            get { return this.payloadClaims[AuthClaimTypes.Audience]; }
        }

        public bool IsValid { get; private set; }

        public string AuthMetadataUri
        {
            get { return this.appContext[AuthClaimTypes.MsExchAuthMetadataUrl]; }
        }

        private void ValidateAudience()
        {
            if (!this.payloadClaims.ContainsKey(AuthClaimTypes.Audience))
            {
                throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", AuthClaimTypes.Audience));
            }


            string appAudience = SettingsHelper.AppBaseUrl + "home/index/";
            string location = appAudience.Replace("/", "-").Replace("\\", "-");
            string audience = this.payloadClaims[AuthClaimTypes.Audience].Replace("/", "-").Replace("\\", "-");

            if (!location.Equals(audience, StringComparison.CurrentCultureIgnoreCase))
            {
                throw new ApplicationException(String.Format(
                  "The audience URL does not match. Expected {0}; got {1}.",
                  appAudience, this.payloadClaims[AuthClaimTypes.Audience]));
            }
        }

        private void ValidateHeaderClaim(string key, string value)
        {
            if (!this.headerClaims.ContainsKey(key))
            {
                throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
            }

            if (!value.Equals(this.headerClaims[key]))
            {
                throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
            }
        }

        private void ValidateHeader()
        {
            ValidateHeaderClaim(AuthClaimTypes.TokenType, Config.TokenType);
            ValidateHeaderClaim(AuthClaimTypes.Algorithm, Config.Algorithm);
        }

        private void ValidateLifetime()
        {
            if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidFrom))
            {
                throw new ApplicationException(
                  String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidFrom));
            }

            if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidTo))
            {
                throw new ApplicationException(
                  String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidTo));
            }

            DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

            TimeSpan padding = new TimeSpan(0, 5, 0);

            DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidFrom]));
            DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidTo]));

            DateTime now = DateTime.UtcNow;

            if (now < (validFrom - padding))
            {
                throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
            }

            if (now > (validTo + padding))
            {
                throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
            }
        }

        private void ValidateMetadataLocation()
        {
            if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchAuthMetadataUrl))
            {
                throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchAuthMetadataUrl));
            }
        }

        private void ValidateVersion()
        {
            if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchTokenVersion))
            {
                throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchTokenVersion));
            }

            if (!Config.Version.Equals(this.appContext[AuthClaimTypes.MsExchTokenVersion]))
            {
                throw new ApplicationException(String.Format(
                  "The version does not match. Expected {0}; got {1}.",
                  Config.Version, this.appContext[AuthClaimTypes.MsExchTokenVersion]));
            }
        }

        #region IDisposable Members

        private bool disposed = false;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                }
            }
            disposed = true;
        }

        #endregion
    }
}