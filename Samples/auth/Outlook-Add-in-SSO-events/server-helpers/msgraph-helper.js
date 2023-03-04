// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the functionality to get Microsoft Graph data. 

const getData = require('./odata-helper');

const domain = 'graph.microsoft.com';
const versionURLsegment = '/v1.0';

/**
 *  Calls a Microsoft Graph API and returns the response.
 *
 * @param {*} accessToken The access token obtained through the On-Behalf_Of flow with correct permissions to Microsoft Graph
 * @param {*} apiURLsegment The Microsoft Graph node to call, e.g., '/me/drive/root/children'
 * @param {*} queryParamsSegment An optional oData query, e.g., '?$select=name&$top=10'
 * @returns
 */
async function getGraphData(accessToken, apiURLsegment, queryParamsSegment) {
  // HTML encodes the parameters to prevent a JavaScript injection attack.
  //  apiURLsegment = encodeURIComponent(apiURLsegment);
  //  queryParamsSegment = encodeURIComponent(queryParamsSegment);

  return new Promise((resolve, reject) => {
    try {
      const oData = getData(
        accessToken,
        domain,
        apiURLsegment,
        versionURLsegment,
        queryParamsSegment
      );
      resolve(oData);
    } catch (error) {
      reject(Error('Unable to call Microsoft Graph. ' + error.toString()));
    }
  });
}

module.exports = getGraphData;
