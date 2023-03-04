// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* 
    This file provides the functionality to get data from OData-compliant endpoints. 
*/

var https = require('https');

const getData = function(accessToken, 
    domain, 
    apiURLsegment, 
    apiVersion, 
    // If any part of queryParamsSegment comes from user input,
    // be sure that it's sanitized so that it can't be used in
    // a Response header injection attack.
    queryParamsSegment) {

    return new Promise((resolve, reject) => {
        var options = {
            host: domain,
            path: apiVersion + apiURLsegment + queryParamsSegment,
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                Accept: 'application/json',
                Authorization: 'Bearer ' + accessToken,
                'Cache-Control': 'private, no-cache, no-store, must-revalidate',
                'Expires': '-1',
                'Pragma': 'no-cache'
            }            
        };

        let req = https.request(options, function (response) {
            var body = '';
            response.on('data', function (d) {
                    body += d;
                });
            response.on('end', function () {

                // The response from the OData endpoint might be an error, say a
                // 401, if the endpoint requires an access token and it was invalid
                // or expired. However, a message isn't an error in the call of https.get,
                // so the "on('error', reject)" line below isn't triggered. 
                // The code distinguishes success (200) messages from error 
                // messages and sends a JSON object to the caller with either the
                // requested OData or error information.

                var error;
                if (response.statusCode === 200) {
                    let parsedBody = JSON.parse(body);
                    resolve(parsedBody);
                } else {
                    error = new Error();
                    error.code = response.statusCode;
                    error.message = response.statusMessage;
                    
                    // The error body sometimes includes an empty space
                    // before the first character. Remove it to avoid causing an error.
                    body = body.trim();
                    error.bodyCode = JSON.parse(body).error.code;
                    error.bodyMessage = JSON.parse(body).error.message;
                    resolve(error);
                }
            });
        })
        .on('error',  reject);
        req.end();
    });
}

module.exports = getData;