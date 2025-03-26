// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const authConfig = {
    credentials: {
        clientID: "Enter_API_Application_Id_Here",
        tenantID: "Enter_API_Tenant_Id_Here",
    },
    metadata: {
        authority: "login.microsoftonline.com",
    },
    protectedRoutes: {
        todolist: {
            endpoint: "/api/todolist",
            delegatedPermissions: {
                read: ["Todolist.Read", "Todolist.ReadWrite"],
                readWrite: ["Todolist.ReadWrite"]
            }
        }
    }
}

module.exports = authConfig;
