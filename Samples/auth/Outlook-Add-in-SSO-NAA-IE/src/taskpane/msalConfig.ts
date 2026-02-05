// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the default MSAL configuration for the add-in project.

import { createLocalUrl } from "./util";

const clientId = "ce1d7062-13da-48b5-940c-564e1f66b535";

export const getMsalConfigShared = () => {
  const msalConfig = {
    auth: {
      clientId,
      redirectUri: createLocalUrl("auth.html"),
      postLogoutRedirectUri: createLocalUrl("auth.html"),
    },
    cache: {
      cacheLocation: "localStorage",
    },
    system: {
      loggerOptions: {},
    },
  };

  return msalConfig;
};

export const defaultScopes = ["user.read", "files.read"]; //CHECK
