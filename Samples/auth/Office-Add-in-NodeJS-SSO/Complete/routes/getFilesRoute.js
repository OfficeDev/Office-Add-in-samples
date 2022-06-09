/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

const express = require("express");
var router = express.Router();
const authHelper = require("../server-helpers/obo-auth-helper");
const getGraphData = require("../server-helpers/msgraph-helper");

router.get(
  "/getuserfilenames",
  authHelper.validateJwt,
  async function (req, res) {
    const authHeader = req.headers.authorization;
    let oboRequest = {
      oboAssertion: authHeader.split(" ")[1],
      scopes: ["files.read.all"],
    };

    const cca = authHelper.getConfidentialClientApplication();
    cca
      .acquireTokenOnBehalfOf(oboRequest)
      .then(async (response) => {
        console.log(response);

        // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
        // and only the top 10 folder or file names.
        const rootUrl = "/me/drive/root/children";
        
        // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
        // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
        // sanitized so that it cannot be used in a Response header injection attack.
        const params = "?$select=name&$top=10";

        let graphData = await getGraphData(
          response.accessToken,
          rootUrl,
          params);
          
            // If Microsoft Graph returns an error, such as invalid or expired token,
            // there will be a code property in the returned object set to a HTTP status (e.g. 401).
            // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.

            if (graphData.code) {
              res
                .status(500)
                .send({ type: "Microsoft Graph", errorDetails: graphData });
            } else {
              // MS Graph data includes OData metadata and eTags that we don't need.
              // Send only what is actually needed to the client: the item names.
              const itemNames = [];
              const oneDriveItems = graphData["value"];
              for (let item of oneDriveItems) {
                itemNames.push(item["name"]);
              }

              res.status(200).send(itemNames);
            }
      })
      .catch((err) => {
        // On rare occasions the SSO access token is unexpired when Office validates it,
        // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
        // with "The provided value for the 'assertion' is not valid. The assertion has expired."
        // Construct an error message to return to the client so it can refresh the SSO token.
        if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
          res.status(500).send({ type: "AADSTS500133", errorDetails: err });
        } else {
          res.status(500).send({ type: err.errorCode, errorDetails: err });
        }
      });
  }
);

module.exports = router;
