// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const jwt = require("jsonwebtoken");
const jwksClient = require("jwks-rsa");
const authConfig = require("./authConfig");

const DISCOVERY_KEYS_ENDPOINT =
 "https://login.microsoftonline.com/common/discovery/v2.0/keys";

// Wrap this with one parameter that returns a new function (req, res, next).
exports.validateJwt = function (req, res, next) {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.split(" ")[1];
    req.token = token;
    const validationOptions = {
      audience: authConfig.credentials.clientID, // v2.0 token      
    };

    jwt.verify(token, getSigningKeys, validationOptions, (err, payload) => {

      if (err) {
        if (err.name === "TokenExpiredError") {
          return res
            .status(401)
            .send({ type: "TokenExpiredError", errorDetails: err });
        } else {
          return res.status(403).send({ type: "Unknown", errorDetails: err });
        }
      }

      // Put claims in the request object for downstream use.
      req.authInfo = payload;
      next();
    });
    

  } else {
    res.status(401).send({ type: "Unknown", errorDetails: "Missing authorization header." });
  }
};

const getSigningKeys = (header, callback) => {
  var client = jwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT,
  });

  client.getSigningKey(header.kid, function (err, key) {
    var signingKey = key.publicKey || key.rsaPublicKey;
    callback(null, signingKey);
  });
};
