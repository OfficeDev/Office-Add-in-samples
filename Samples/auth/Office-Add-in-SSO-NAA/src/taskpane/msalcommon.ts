// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides common MSAL functions for use in the add-in project.

import {
  type RedirectRequest,
} from "@azure/msal-browser";

// Constants
const PROMPT_SELECT_ACCOUNT = "select_account" as const;

/**
 * Creates a token request object with the specified parameters.
 * @param scopes - Array of OAuth 2.0 scopes to request access for.
 * @param selectAccount - Whether to prompt for account selection.
 * @param redirectUri - Optional redirect URI for the authentication flow.
 * @param loginHint - Optional login hint to pre-populate the username field.
 * @returns A properly configured RedirectRequest object.
 * @throws {Error} When scopes array is empty or invalid.
 */
export function getTokenRequest(
  scopes: string[], 
  selectAccount: boolean, 
  redirectUri?: string, 
  loginHint?: string
): RedirectRequest {
  // Validate required parameters.
  if (!scopes || scopes.length === 0) {
    throw new Error("Scopes array cannot be empty");
  }

  // Build request object.
  const request: RedirectRequest = { scopes };

  if (loginHint) {
    request.loginHint = loginHint;
  }

  if (selectAccount) {
    request.prompt = PROMPT_SELECT_ACCOUNT;
  }

  if (redirectUri) {
    request.redirectUri = redirectUri;
  }

  return request;
}

