/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */

import fetch from "node-fetch";
import form from "form-urlencoded";
import jwt from "jsonwebtoken";
import { JwksClient } from "jwks-rsa";

/* global process, console */

const DISCOVERY_KEYS_ENDPOINT = "https://login.microsoftonline.com/common/discovery/v2.0/keys";

export async function getAccessToken(authorization: string): Promise<any> {
  if (!authorization) {
    let error = new Error("No Authorization header was found.");
    return Promise.reject(error);
  } else {
    const scopeName: string = process.env.SCOPE || "User.Read";
    const [, /* schema */ assertion] = authorization.split(" ");

    const tokenScopes = (jwt.decode(assertion) as jwt.JwtPayload).scp.split(" ");
    const accessAsUserScope = tokenScopes.find((scope) => scope === "access_as_user");
    if (!accessAsUserScope) {
      throw new Error("Missing access_as_user");
    }

    const formParams = {
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: assertion,
      requested_token_use: "on_behalf_of",
      scope: [scopeName].join(" "),
    };

    const stsDomain: string = "https://login.microsoftonline.com";
    const tenant: string = "common";
    const tokenURLSegment: string = "oauth2/v2.0/token";
    const encodedForm = form(formParams);

    const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
      method: "POST",
      body: encodedForm,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });
    const json = await tokenResponse.json();
    return json;
  }
}

export function validateJwt(req, res, next): void {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.split(" ")[1];

    const validationOptions = {
      audience: process.env.CLIENT_ID,
    };

    jwt.verify(token, getSigningKeys, validationOptions, (err) => {
      if (err) {
        console.log(err);
        return res.sendStatus(403);
      }

      next();
    });
  }
}

function getSigningKeys(header: any, callback: any) {
  var client: JwksClient = new JwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT,
  });

  client.getSigningKey(header.kid, function (err, key) {
    callback(null, key.getPublicKey());
  });
}
