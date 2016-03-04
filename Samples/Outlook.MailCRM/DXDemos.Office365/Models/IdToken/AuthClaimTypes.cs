using System;
using Microsoft.IdentityModel.S2S.Tokens;

namespace DXDemos.Office365.IdToken.Models
{
    /// <summary>
    /// Everything in this folder is an exact reference from https://msdn.microsoft.com/en-us/library/office/fp179819.aspx on how to validate an Exchange identity token
    /// </summary>
    public class AuthClaimTypes
    {
        public const string NameIdentifier =
            JsonWebTokenConstants.ReservedClaims.NameIdentifier;
        public const string MsExchImmutableId = "msexchuid";
        public const string MsExchTokenVersion = "version";
        public const string MsExchAuthMetadataUrl = "amurl";

        public const string AppContext =
            JsonWebTokenConstants.ReservedClaims.AppContext;
        public const string Audience =
            JsonWebTokenConstants.ReservedClaims.Audience;
        public const string Issuer =
            JsonWebTokenConstants.ReservedClaims.Issuer;
        public const string ValidFrom =
            JsonWebTokenConstants.ReservedClaims.NotBefore;
        public const string ValidTo =
            JsonWebTokenConstants.ReservedClaims.ExpiresOn;

        public const string AppContextSender = "appctxsender";
        public const string IsBrowserHostedApp = "isbrowserhostedapp";

        public const string TokenType = "typ";
        public const string Algorithm = "alg";
        public const string x509Thumbprint = "x5t";
    }
}