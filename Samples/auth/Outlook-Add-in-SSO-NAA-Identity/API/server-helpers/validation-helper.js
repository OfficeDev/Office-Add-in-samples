// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const jwt = require("jsonwebtoken");
const jwksClient = require("jwks-rsa");
const authConfig = require("./authConfig");

const DISCOVERY_KEYS_ENDPOINT =
  "https://login.microsoftonline.com/common/discovery/v2.0/keys";

// Wrap this with one parameter that returns a new function (req, res, next).
exports.validateJwt = async function (req, res, next) {
  try {
    const authHeader = req.headers.authorization;
    if (authHeader) {
      const token = authHeader.split(" ")[1];
      req.token = token;

      // Validate issuer (see https://learn.microsoft.com/entra/identity-platform/access-tokens#multitenant-applications)
      const decodedToken = jwt.decode(token, { complete: true });
      const iss = decodedToken.payload.iss;
      const response = await fetch('https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration');
      const openidConfiguration = await response.json();
      
      // Replace the placeholder with the actual tenant ID.
      const expectedIssuer = openidConfiguration.issuer.replace("{tenantid}",authConfig.credentials.tenantID);
      if (iss !== expectedIssuer) {
        return res.status(401).send({ type: "Unknown", errorDetails: "Invalid issuer." });
      }

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
  } catch (error) {
    console.error("Error in validateJwt:", error);
    res.status(500).send({ type: "Unknown", errorDetails: "Internal server error." });
  };
}

const getSigningKeys = (header, callback) => {
  var client = jwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT,
  });

  client.getSigningKey(header.kid, function (err, key) {
    var signingKey = key.publicKey || key.rsaPublicKey;
    callback(null, signingKey);
  });
};
