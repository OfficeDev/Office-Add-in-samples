// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const msal = require("@azure/msal-node");
const jwt = require("jsonwebtoken");
const jwksClient = require("jwks-rsa");

const DISCOVERY_KEYS_ENDPOINT =
  "https://login.microsoftonline.com/" +
  process.env.DIRECTORY_ID +
  "/discovery/v2.0/keys";

const config = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: "https://login.microsoftonline.com/" + process.env.DIRECTORY_ID,
    clientSecret: process.env.CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message, containsPii) {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Verbose,
    },
  },
};

exports.getConfidentialClientApplication = function getConfidentialClientApplication(){
  // Create msal application object
  return new msal.ConfidentialClientApplication(config);
}

exports.validateJwt = function (req, res, next) {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.split(" ")[1];

    const validationOptions = {
      audience: config.auth.clientId, // v2.0 token
      issuer: config.auth.authority + "/v2.0", // v2.0 token
    };

    jwt.verify(token, getSigningKeys, validationOptions, (err, payload) => {
      if (err) {
        console.log(err);
        return res.sendStatus(403);
      }

      next();
    });
  } else {
    res.sendStatus(401);
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
