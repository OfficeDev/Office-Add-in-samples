// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the provides functionality to get Microsoft Graph data.

/* global console fetch */

import { AccountManager } from "./authConfig";

/**
 *  Calls a Microsoft Graph API and returns the response.
 *
 * @param accessToken The access token to use for the request.
 * @param path Path component of the URI, e.g., "/me". Should start with "/".
 * @param queryParams Query parameters, e.g., "?$select=name,id". Should start with "?".
 * @returns
 */
export async function makeGraphRequest(accountManager: AccountManager, scopes: string[], path: string, queryParams: string): Promise<any> {
  if (!path) throw new Error("path is required.");
  if (!path.startsWith("/")) throw new Error("path must start with '/'.");
  if (queryParams && !queryParams.startsWith("?")) throw new Error("queryParams must start with '?'.");

  // Get the access token
  let accessToken = await accountManager.ssoGetAccessToken(scopes);

  const response = await fetch(`https://graph.microsoft.com/v1.0${path}${queryParams}`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    const data = await response.json();
    console.log(data);
    return data;
  } else {
    // Check for CAE claims challenge.
    if (response.status === 401 && response.headers.get('www-authenticate')) {
      const authenticateHeader: string = response.headers.get('www-authenticate') as string;
      const claimsChallenge = parseChallenges(authenticateHeader).claims;
      // use the claims challenge to acquire a new access token...
      accessToken = await accountManager.ssoGetAccessToken(scopes, claimsChallenge);
      // Attempt the MS Graph call again.
      const response2 = await fetch(`https://graph.microsoft.com/v1.0${path}${queryParams}`, {
        headers: { Authorization: accessToken },
      });

      if (response2.ok) {
        const data = await response2.json();
        console.log(data);
        return data;
      } else {
        // Still not successful. Throw the error.
        throw new Error(response2.statusText);
      }
    } else {
      throw new Error(response.statusText);
    }
  }
}

// helper function to parse the www-authenticate header
function parseChallenges(header: string): { [key: string]: string } {
  const schemeSeparator = header.indexOf(' ');
  const challenges = header.substring(schemeSeparator + 1).split(',');
  const challengeMap: { [key: string]: string } = {};

  challenges.forEach((challenge) => {
    const [key, value] = challenge.split('=');
    challengeMap[key.trim()] = window.decodeURI(value.replace(/['"]+/g, ''));
  });
  return challengeMap;
}
