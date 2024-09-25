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

/**
 * Gets a token request for a given account context.
 * @param accountContext The account context to get the token request for.
 * @returns The token request.
 */
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

/**
 * Returns the existing public client application. Returns a new public client application if it did not exist.
 * @returns The nested public client application.
 */
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

/**
 * Gets the account information of the given user.
 * @param accountContext The account context of the user. If not provided the function gets the active account.
 * @returns The account information of the user.
 */
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
