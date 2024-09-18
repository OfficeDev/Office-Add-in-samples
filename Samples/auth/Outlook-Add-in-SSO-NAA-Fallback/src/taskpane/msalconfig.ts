import { createLocalUrl } from "./util";

export const clientId = "b3b6dc33-016b-4bf7-b4b6-c97fa56c1879";
export const msalConfig = {
  auth: {
    clientId,
    redirectUri: createLocalUrl("auth.html"),
    postLogoutRedirectUri: createLocalUrl("auth.html"),
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        console.log(message);
      },
    },
  },
};

export const defaultScopes = ["user.read"];
