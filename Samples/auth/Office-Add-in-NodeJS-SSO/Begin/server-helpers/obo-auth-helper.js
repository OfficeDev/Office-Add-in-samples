// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const msal = require("@azure/msal-node");
const jwt = require("jsonwebtoken");
const jwksClient = require("jwks-rsa");

const DISCOVERY_KEYS_ENDPOINT =
  "https://login.microsoftonline.com/common/discovery/v2.0/keys";

const config = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: "https://login.microsoftonline.com/common",
    clientSecret: process.env.CLIENT_SECRET,
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

    const validationOptions = {
      audience: config.auth.clientId, // v2.0 token
      //issuer: config.auth.authority + "/v2.0", // v2.0 token  **can't use this one
    };

    jwt.verify(token, getSigningKeys, validationOptions, (err, payload) => {
      //custom logic to regex search for tenant id in the issuer.
      //test multi tenant setup.
      //test msa

      if (err) {
        // On rare occasions the SSO access token is unexpired when Office validates it,
        // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
        // with "The provided value for the 'assertion' is not valid. The assertion has expired."
        // Construct an error message to return to the client so it can refresh the SSO token.
        if (err.name === "TokenExpiredError") {
          return res
            .status(401)
            .send({ type: "TokenExpiredError", errorDetails: err });
        } else {
          return res.status(403).send({ type: "Unknown", errorDetails: err });
        }
      }
      next();
    });
  } else {
    res.status(401).send({ type: "Unknown", errorDetails: err });
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
