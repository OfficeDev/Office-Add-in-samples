// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const msal = require("@azure/msal-node");
const jwt = require("jsonwebtoken");
const jwksClient = require("jwks-rsa");

const DISCOVERY_KEYS_ENDPOINT =
  "https://login.microsoftonline.com/common/discovery/v2.0/keys";

const config = {
  auth: {
    clientId: "d0ff5b49-5a37-4b63-889e-86cf8821eb07", // client id from API app registration
    authority: "https://login.microsoftonline.com/common",
    //clientSecret: process.env.CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message, containsPii) {
        if (containsPii) {
          return;
        }
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Verbose,
    },
  },
};

exports.getConfidentialClientApplication =
  function getConfidentialClientApplication() {
    // Create msal application object
    return new msal.ConfidentialClientApplication(config);
  };

// wrap this with one parameter that returns a new function (req,res,next)
exports.validateJwt = function (req, res, next) {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.split(" ")[1];
    req.token = token;
    const validationOptions = {
      audience: config.auth.clientId, // v2.0 token
      //issuer: config.auth.authority + "/v2.0", // v2.0 token  **can't use this one
    };

    
    jwt.verify(token, getSigningKeys, validationOptions, (err, payload) => {
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
