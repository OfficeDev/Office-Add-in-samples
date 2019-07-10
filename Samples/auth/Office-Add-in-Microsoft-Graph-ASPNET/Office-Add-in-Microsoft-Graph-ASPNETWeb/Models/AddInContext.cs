// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System.Data.Entity;

namespace OfficeAddinMicrosoftGraphASPNET.Models
{
    /// <summary>
    /// Models the session tokens in the database.
    /// </summary>
    public class AddInContext : DbContext
    {
        public AddInContext() : base("AddInContext")
        {
        }

        public DbSet<SessionToken> SessionTokens { get; set; }
    }
}
