// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the provides functionality to get Microsoft Graph data.

/* global console fetch */

//const getData = require("./odata-helper");

export { getGraphData };

// const domain = "graph.microsoft.com";
// const versionURLsegment = "/v1.0";

/**
 *  Calls a Microsoft Graph API and returns the response.
 *
 * @param {*} accessToken The access token obtained through the On-Behalf_Of flow with correct permissions to Microsoft Graph
 * @param {*} apiURLsegment The Microsoft Graph node to call, e.g., "/me/drive/root/children"
 * @param {*} queryParamsSegment An optional oData query, e.g., "?$select=name&$top=10"
 * @returns
 */
async function getGraphData(accessToken: string, apiURLsegment: string, queryParamsSegment: string) {
  const requestString = "https://graph.microsoft.com/v1.0" + apiURLsegment + queryParamsSegment;
  const headersInit = { Authorization: accessToken };
  const requestInit = { headers: headersInit };
  const result = await fetch(requestString, requestInit);
  if (!result.ok) {
    // error
    throw new Error(result.statusText);
  }
  const data = await result.json();
  console.log(data);
  return data;
}
