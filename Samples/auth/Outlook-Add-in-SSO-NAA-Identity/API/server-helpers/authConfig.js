const authConfig = {
    credentials: {
        tenantID: "255cd1be-82ef-4e95-9fb2-a2d9b705151f",
        clientID: "d0ff5b49-5a37-4b63-889e-86cf8821eb07"
    },
    metadata: {
        authority: "login.microsoftonline.com",
        discovery: ".well-known/openid-configuration",
        version: "v2.0"
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
