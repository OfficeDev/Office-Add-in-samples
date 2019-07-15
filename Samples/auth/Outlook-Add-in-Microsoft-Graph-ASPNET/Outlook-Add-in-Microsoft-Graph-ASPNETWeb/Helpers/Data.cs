// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using OutlookAddinMicrosoftGraphASPNET.Models;
using System.Linq;

namespace OutlookAddinMicrosoftGraphASPNET.Helpers
{
    /// <summary>
    /// Handles session ID storage.
    /// </summary>
    public static class Data
    {
        /// <summary>
        /// Gets the user session token from the database.
        /// </summary>
        /// <param name="userAuthSessionId"></param>
        /// <param name="provider"></param>
        /// <returns></returns>
        public static SessionToken GetUserSessionToken(string userAuthSessionId, string provider)
        {
            SessionToken st = null;
            using (var db = new AddInContext())
            {
                st = db.SessionTokens.FirstOrDefault(t => t.Id == userAuthSessionId && t.Provider == provider);
            }
            return st;
        }

        public static void DeleteUserSessionToken(string userAuthSessionId, string provider)
        {
            using (var db = new AddInContext())
            {
                var st = db.SessionTokens.Where(t => t.Id == userAuthSessionId && t.Provider == provider);
                if (st.Any())
                {
                    db.SessionTokens.RemoveRange(st);
                    db.SaveChanges();
                }
            }
        }
    }
}
