// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const express = require('express');
var router = express.Router();
const authHelper = require('../server-helpers/obo-auth-helper');
const jwt = require('jsonwebtoken');
const { v4: uuidv4 } = require('uuid');
const authConfig = require('../authConfig');


router.get('/todolist', authHelper.validateJwt, getTodos);

//router.get('/todolist/:id', authHelper.validateJwt, todolist.getTodo);

router.post('/todolist', authHelper.validateJwt, postTodo);

function getTodos(req, res, next){
  res.status(200).send("testss");
}

function postTodo(req, res, next) {
  
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.write)
        ||
        hasRequiredApplicationPermissions(req.authInfo, authConfig.protectedRoutes.todolist.applicationPermissions.write)
    ) {
        try {
            const todo = {
                description: req.body.description,
                id: uuidv4(),
                owner: req.authInfo['oid'] // oid is the only claim that should be used to uniquely identify a user in an Azure AD tenant
            };

            db.get('todos').push(todo).write();

            res.status(200).json(todo);
        } catch (error) {
            next(error);
        }
    } else (
        next(new Error('User or application does not have the required permissions'))
    )
}

/**
 * Ensures that the access token has the specified delegated permissions.
 * @param {Object} accessTokenPayload: Parsed access token payload
 * @param {Array} requiredPermission: list of required permissions
 * @returns {boolean}
 */
const hasRequiredDelegatedPermissions = (accessTokenPayload, requiredPermission) => {
  const normalizedRequiredPermissions = requiredPermission.map(permission => permission.toUpperCase());

  if (accessTokenPayload.hasOwnProperty('scp') && accessTokenPayload.scp.split(' ')
      .some(claim => normalizedRequiredPermissions.includes(claim.toUpperCase()))) {
      return true;
  }

  return false;
}

router.get(
  '/getuserfilenames',
  authHelper.validateJwt,
  async function (req, res) {
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(' ')[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(' ');
      let accessAsUserScope = tokenScopes.find(
        (scope) => scope === 'access_as_user'
      );
      if (accessAsUserScope !== 'access_as_user' ) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
      // and only the top 10 folder or file names.
      const rootUrl = '/me/drive/root/children';

      // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
      // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
      // sanitized so that it cannot be used in a Response header injection attack.
      const params = '?$select=name&$top=10';

      const graphData = await getGraphData(
        response.accessToken,
        rootUrl,
        params
      );

      // If Microsoft Graph returns an error, such as invalid or expired token,
      // there will be a code property in the returned object set to a HTTP status (e.g. 401).
      // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
      if (graphData.code) {
        res
          .status(403)
          .send({
            type: "Microsoft Graph",
            errorDetails:
              "An error occurred while calling the Microsoft Graph API.\n" +
              graphData,
          });
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
    } catch (err) {
      // On rare occasions the SSO access token is unexpired when Office validates it,
      // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
      // with "The provided value for the 'assertion' is not valid. The assertion has expired."
      // Construct an error message to return to the client so it can refresh the SSO token.
      if ((err.errorMessage !== undefined) && err.errorMessage.indexOf('AADSTS500133') !== -1) {
        res.status(401).send({ type: "TokenExpiredError", errorDetails: JSON.stringify(err) });
      } else {
        res.status(403).send({ type: "Unknown", errorDetails: JSON.stringify(err) });
      }
    }
  }
);

module.exports = router;
