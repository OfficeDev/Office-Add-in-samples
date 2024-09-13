import { createLocalUrl } from "./util";

export const clientId = "148b0448-c6ab-4d8e-adb2-a0f2696966d2";
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
