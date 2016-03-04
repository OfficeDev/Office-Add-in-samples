using System;
using System.Text;

namespace DXDemos.Office365.IdToken.Models
{
    /// <summary>
    /// Everything in this folder is an exact reference from https://msdn.microsoft.com/en-us/library/office/fp179819.aspx on how to validate an Exchange identity token
    /// </summary>
    public static class Base64UrlEncoder
    {
        public static Encoding TextEncoding = Encoding.UTF8;

        private static char Base64PadCharacter = '=';
        private static char Base64Character62 = '+';
        private static char Base64Character63 = '/';
        private static char Base64UrlCharacter62 = '-';
        private static char Base64UrlCharacter63 = '_';

        private static byte[] DecodeBytes(string arg)
        {
            if (String.IsNullOrEmpty(arg))
            {
                throw new ApplicationException("String to decode cannot be null or empty.");
            }

            StringBuilder s = new StringBuilder(arg);
            s.Replace(Base64UrlCharacter62, Base64Character62);
            s.Replace(Base64UrlCharacter63, Base64Character63);

            int pad = s.Length % 4;
            s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

            return Convert.FromBase64String(s.ToString());
        }

        public static string Decode(string arg)
        {
            return TextEncoding.GetString(DecodeBytes(arg));
        }
    }
}