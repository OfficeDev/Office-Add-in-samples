// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the default MSAL configuration for the add-in project.

import { createLocalUrl } from "./util";

/* global console */

const clientId = "b3b6dc33-016b-4bf7-b4b6-c97fa56c1879";

// CHECK
/**
 * Log message level generic for local use (copied from MSAL Logger.ts to avoid importing MSAL v3).
 */
export enum LogLevelLocal {
  Error,
  Warning,
  Info,
  Verbose,
  Trace,
}

export const getMsalConfig = (enableDebugLogging: boolean) => {
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
  if (enableDebugLogging && msalConfig.system) {
    msalConfig.system.loggerOptions = {
      loggerOptions: {
        logLevel: LogLevelLocal.Verbose,
        loggerCallback: (level: LogLevelLocal, message: string) => {
          switch (level) {
            case LogLevelLocal.Error:
              console.error(message);
              return;
            case LogLevelLocal.Info:
              console.info(message);
              return;
            case LogLevelLocal.Verbose:
              console.debug(message);
              return;
            case LogLevelLocal.Warning:
              console.warn(message);
              return;
          }
        },
        piiLoggingEnabled: true,
      },
    };
  }
  return msalConfig;
};

export const defaultScopes = ["user.read", "files.read"]; //CHECK
