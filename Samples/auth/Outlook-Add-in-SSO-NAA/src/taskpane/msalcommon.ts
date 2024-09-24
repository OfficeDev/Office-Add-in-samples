// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides common MSAL functions for use in the add-in project.

import {
  AccountInfo,
  createNestablePublicClientApplication,
  PublicClientApplication,
  type RedirectRequest,
} from "@azure/msal-browser";
import { defaultScopes, msalConfig } from "./msalconfig";

export async function getTokenRequest(accountContext?: AccountContext): Promise<RedirectRequest> {
  const account = await getAccountFromContext(accountContext);
  let additionalProperties: Partial<RedirectRequest> = {};
  if (account) {
    additionalProperties = { account };
  } else if (accountContext) {
    additionalProperties = {
      loginHint: accountContext.loginHint,
    };
  } else {
    additionalProperties = { prompt: "select_account" };
  }
  return { scopes: defaultScopes, ...additionalProperties };
}

let _publicClientApp: PublicClientApplication;

export async function ensurePublicClient() {
  if (!_publicClientApp) {
    _publicClientApp = await createNestablePublicClientApplication(msalConfig);
  }
  return _publicClientApp;
}

export type AccountContext = {
  loginHint?: string;
  tenantId?: string;
  localAccountId?: string;
};

export async function getAccountFromContext(accountContext?: AccountContext): Promise<AccountInfo | null> {
  const pca = await ensurePublicClient();
  if (!accountContext) {
    return pca.getActiveAccount();
  }

  return pca.getAccount({
    username: accountContext.loginHint,
    tenantId: accountContext.tenantId,
    localAccountId: accountContext.localAccountId,
  });
}
