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
    const oboRequest = {
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
        // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
        // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
        // sanitized so that it cannot be used in a Response header injection attack.
        const graphData = await getGraphData(
          response.accessToken,
          "/me/drive/root/children",
          "?$select=name&$top=10"
        );

        // If Microsoft Graph returns an error, such as invalid or expired token,
        // there will be a code property in the returned object set to a HTTP status (e.g. 401).
        // Relay it to the client. It will caught in the fail callback of `makeGraphApiCall`.
        if (graphData.code) {
          next(
            createError(
              graphData.code,
              "Microsoft Graph error " + JSON.stringify(graphData)
            )
          );
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
      .catch((error) => {
        // Handle request for multi-factor authentication.
        if (error.Message.StartsWith("AADSTS50076")) {
          responseMessage = String.Format(
            '{{"AADError":"AADSTS50076","Claims":{0}}}',
            error.Claims
          );
          res.status(403).send(responseMessage);
          // The client should recall the getAccessToken function and pass the claims string as the
          // authChallenge value in the function's Options parameter.
        }

        // Handle lack of consent (AADSTS65001) and invalid scope (permission).
        if (
          error.Message.StartsWith("AADSTS65001") ||
          error.Message.StartsWith(
            "AADSTS70011: The provided value for the input parameter 'scope' is not valid."
          )
        ) {
          res.status(403).send(error);
          return HttpErrorHelper.SendErrorToClient(
            HttpStatusCode.Forbidden,
            e,
            null
          );
        }

        // Handle all other MsalServiceExceptions.
        res.status(500).send(error);
      });
  }
);

module.exports = router;
