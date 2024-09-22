// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Represents a user profile from an MSAL account.
 */
export interface UserProfile {
  userName: string;
  userEmail: string;
  accessToken: string;
}

export enum AuthMethod {
  NAA,
  MSALV3, // For when NAA unavailable.
  MSALV2, // For Trident IE11 webview support.
}
