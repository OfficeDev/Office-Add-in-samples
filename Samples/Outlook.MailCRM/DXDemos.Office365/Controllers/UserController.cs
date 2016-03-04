using DXDemos.Office365.IdToken.Models;
using DXDemos.Office365.Models;
using DXDemos.Office365.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace DXDemos.Office365.Controllers
{
    public class UserController : ApiController
    {
        [ActionName("DefaultAction")]
        public UserModel Get(string id)
        {
            return DocumentDBRepository<UserModel>.GetItem("Users", i => i.id == id);
        }

        // Post: api/User/Validate/
        [HttpPost]
        [Route("api/User/Validate/")]
        public async Task<HttpResponseMessage> Validate([FromBody]IdentityTokenRequest token)
        {
            //validate the identity token passed from the client
            IdentityTokenResponse response = new IdentityTokenResponse();

            try
            {
                //decode and validate the token passed in
                IdentityToken identityToken = null;
                using (DecodedJsonToken decodedToken = JsonTokenDecoder.Decode(token))
                {
                    if (decodedToken.IsValid)
                    {
                        identityToken = new IdentityToken(token, decodedToken.Audience, decodedToken.AuthMetadataUri);
                    }
                }
                response.token = identityToken;

                //now that the key is validated, we can perform a lookup against DocDB for it's hased value (combination of metadata document URL with the Exchange identifier)
                if (identityToken != null)
                {
                    //the token is valid...check if user is valid (has valid refresh token)
                    response.validToken = true;
                    string hash = ComputeSHA256Hash(response.token.uniqueID, response.token.amurl, Salt);
                    response.user = DocumentDBRepository<UserModel>.GetItem("Users", i => i.hash == hash);
                    if (response.user != null)
                    {
                        //check for and validate the refresh token
                        if (!String.IsNullOrEmpty(response.user.refresh_token)) 
                        {
                            var graphToken = await TokenHelper.GetAccessTokenWithRefreshToken(response.user.refresh_token, SettingsHelper.O365UnifiedAPIResourceId);
                            if (graphToken != null)
                            {
                                //TODO: get the user details against AAD Graph
                                response.validUser = true;
                            }
                        }

                    }
                    else 
                    {
                        //the user doesn't exist, so we can add a placeholder record in the data store...TODO: get more data on user????
                        response.user = new UserModel() { id = Guid.NewGuid().ToString().ToLower(), hash = hash };
                        await DocumentDBRepository<UserModel>.CreateItemAsync("Users", response.user);
                    }
                }
                else
                {
                    //this was an invalid token!!!!
                }
            }
            catch (Exception ex)
            {
                response.errorMessage = ex.Message;
            }

            return Request.CreateResponse<IdentityTokenResponse>(HttpStatusCode.OK, response);
        }

        internal static string getInititials(string name, string email) 
        {
            if (!String.IsNullOrEmpty(name))
            {
                var names = name.Split(' ');
                if (names.Length >= 2)
                    return (names[0].Substring(0, 1) + names[1].Substring(0, 1)).ToUpper();
                else
                    return name.Substring(0, 2).ToUpper();
            }
            else
                return email.Substring(0, 2).ToUpper();
        }


        private byte[] Salt = new byte[] { 25, 139, 201, 13 };
        private string ComputeSHA256Hash(string uniqueId, string authenticationMetadataUrl, byte[] salt)
        {
            byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(uniqueId, authenticationMetadataUrl));

            // Combine input bytes and salt.
            byte[] saltedInput = new byte[salt.Length + inputBytes.Length];
            salt.CopyTo(saltedInput, 0);
            inputBytes.CopyTo(saltedInput, salt.Length);

            // Compute the unique key.
            byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

            // Convert the hashed value to a string and return.
            return BitConverter.ToString(hashedBytes);
        }
    }
}
