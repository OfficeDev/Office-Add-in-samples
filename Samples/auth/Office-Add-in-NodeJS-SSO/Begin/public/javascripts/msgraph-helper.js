// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/* 
    This file provides the provides functionality to get Microsoft Graph data. 
*/

var getData = require('./odata-helper');

let domain = "graph.microsoft.com";
let versionURLsegment = "/v1.0";

// If any part of queryParamsSegment comes from user input,
// be sure that it is sanitized so that it cannot be used in
// a Response header injection attack.

async function getGraphData(accessToken, apiURLsegment, queryParamsSegment) {

    return new Promise(async (resolve, reject) => { 

        try {
            const oData = await getData(accessToken, domain, apiURLsegment, versionURLsegment, queryParamsSegment);
            resolve(oData);
        }
        catch(error) {
            reject(Error("Unable to call Microsoft Graph. " + error.toString()));
        }
    })        
} 

module.exports = getGraphData;