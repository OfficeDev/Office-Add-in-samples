// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.IdentityModel.Tokens;

namespace OutlookAddinMicrosoftGraphASPNET.Models
{
    /// <summary>
    /// Models a row of the SessionToken table in the database.
    /// </summary>
    public class SessionToken
    {
        /// <summary>
        /// This is the user SessionID
        /// </summary>
        [Key, Column(Order = 1)]
        [MaxLength(36)]
        public string Id { get; set; }

        // The user identity provider
        [Key, Column(Order = 2)]
        [MaxLength(150)]
        public string Provider { get; set; }

        // The access token for the OData endpoint.
        public string AccessToken { get; set; }

        public DateTime CreatedOn { get; set; }

        [MaxLength(100)]
        public string Username { get; set; }

        //TODO: Validate the token so we can extract the user name and user id properties from the id_token
        public static JwtSecurityToken ParseJwtToken(string jwtToken)
        {
            JwtSecurityToken jst = new JwtSecurityToken(jwtToken);
            return jst;
        }
    }
}
