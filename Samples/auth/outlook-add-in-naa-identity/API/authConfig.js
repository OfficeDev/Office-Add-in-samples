const passportConfig = {
    credentials: {
        tenantID: "Enter_the_Tenant_Info_Here",
        clientID: "Enter_the_Application_Id_Here"
    },
    metadata: {
        authority: "login.microsoftonline.com",
        discovery: ".well-known/openid-configuration",
        version: "v2.0"
    },
    settings: {
        validateIssuer: true,
        passReqToCallback: true,
        loggingLevel: "info",
        loggingNoPII: true,
    },
    protectedRoutes: {
        todolist: {
            endpoint: "/api/todolist",
            delegatedPermissions: {
                read: ["Todolist.Read", "Todolist.ReadWrite"],
                write: ["Todolist.ReadWrite"]
            },
            applicationPermissions: {
                read: ["Todolist.Read.All", "Todolist.ReadWrite.All"],
                write: ["Todolist.ReadWrite.All"]
            }
        }
    }
}

module.exports = passportConfig;