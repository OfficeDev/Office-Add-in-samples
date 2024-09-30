// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the default MSAL configuration for the add-in project.

import { LogLevel } from "@azure/msal-browser";
import { defaultScopes, getMsalConfigShared } from "./msalConfig";

/* global console */

export const getMsalConfig = (enableDebugLogging: boolean) => {
  const msalConfig = getMsalConfigShared();
  if (enableDebugLogging && msalConfig.system) {
    msalConfig.system.loggerOptions = {
      loggerOptions: {
        logLevel: LogLevel.Verbose,
        loggerCallback: (level: LogLevel, message: string) => {
          switch (level) {
            case LogLevel.Error:
              console.error(message);
              return;
            case LogLevel.Info:
              console.info(message);
              return;
            case LogLevel.Verbose:
              console.debug(message);
              return;
            case LogLevel.Warning:
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

export { defaultScopes };
