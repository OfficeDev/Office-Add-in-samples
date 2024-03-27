// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the provides functionality to get Microsoft Graph data.

/* global console fetch */

/**
 *  Calls a Microsoft Graph API and returns the response.
 *
 * @param accessToken The access token to use for the request.
 * @param path Path component of the URI, e.g., "/me". Should start with "/".
 * @param queryParams Query parameters, e.g., "?$select=name,id". Should start with "?".
 * @returns
 */
export async function makeGraphRequest(accessToken: string, path: string, queryParams: string): Promise<any> {
  if (!path) throw new Error("path is required.");
  if (!path.startsWith("/")) throw new Error("path must start with '/'.");
  if (queryParams && !queryParams.startsWith("?")) throw new Error("queryParams must start with '?'.");

  const response = await fetch(`https://graph.microsoft.com/v1.0${path}${queryParams}`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    const data = await response.json();
    console.log(data);
    return data;
  } else {
    throw new Error(response.statusText);
  }
}
