/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

/**
 * This method calls the Graph API by utilizing the graph client instance.
 * @param {String} username 
 * @param {Array} scopes 
 * @param {String} uri 
 * @param {String} interactionType 
 * @param {Object} myMSALObj 
 * @returns 
 */
const callGraph = async (username, scopes, uri, interactionType, myMSALObj) => {
    const account = myMSALObj.getAccountByUsername(username);
    try {
        let response = await getGraphClient({
            account: account,
            scopes: scopes,
            interactionType: interactionType,
        })
            .api(uri)
            .responseType('raw')
            .get();

        response = await handleClaimsChallenge(account, response, uri);
        if (response && response.error === 'claims_challenge_occurred') throw response.error;
        updateUI(response, uri);
    } catch (error) {
        if (error === 'claims_challenge_occurred') {
            const resource = new URL(uri).hostname;
            const claims =
                account &&
                    getClaimsFromStorage(`cc.${msalConfig.auth.clientId}.${account.idTokenClaims.oid}.${resource}`)
                    ? window.atob(
                        getClaimsFromStorage(
                            `cc.${msalConfig.auth.clientId}.${account.idTokenClaims.oid}.${resource}`
                        )
                    )
                    : undefined; // e.g {"access_token":{"xms_cc":{"values":["cp1"]}}}
            let request = {
                account: account,
                scopes: scopes,
                claims: claims,
            };
            switch (interactionType) {
                case msal.InteractionType.Popup:

                    await myMSALObj.acquireTokenPopup({
                        ...request,
                        redirectUri: '/redirect',
                    });
                    break;
                case msal.InteractionType.Redirect:
                    await myMSALObj.acquireTokenRedirect(request);
                    break;
                default:
                    await myMSALObj.acquireTokenRedirect(request);
                    break;
            }
        } else if (error.toString().includes('404')) {
            return updateUI(null, uri);
        } else {
            console.log(error);
        }
    }
}

/**
 * This method inspects the HTTPS response from a fetch call for the "www-authenticate header"
 * If present, it grabs the claims challenge from the header and store it in the localStorage
 * For more information, visit: https://docs.microsoft.com/en-us/azure/active-directory/develop/claims-challenge#claims-challenge-header-format
 * @param {object} response
 * @returns response
 */
const handleClaimsChallenge = async (account, response, apiEndpoint) => {
    if (response.status === 200) {
        return response.json();
    } else if (response.status === 401) {
        if (response.headers.get('WWW-Authenticate')) {
            const authenticateHeader = response.headers.get('WWW-Authenticate');
            const claimsChallenge = parseChallenges(authenticateHeader);
            /**
             * This method stores the claim challenge to the session storage in the browser to be used when acquiring a token.
             * To ensure that we are fetching the correct claim from the storage, we are using the clientId
             * of the application and oid (user’s object id) as the key identifier of the claim with schema
             * cc.<clientId>.<oid>.<resource.hostname>
             */
            addClaimsToStorage(
                claimsChallenge.claims,
                `cc.${msalConfig.auth.clientId}.${account.idTokenClaims.oid}.${new URL(apiEndpoint).hostname}`
            );
            return { error: 'claims_challenge_occurred', payload: claimsChallenge.claims };
        }

        throw new Error(`Unauthorized: ${response.status}`);
    } else {
        throw new Error(`Something went wrong with the request: ${response.status}`);
    }
};

/**
 * This method parses WWW-Authenticate authentication headers
 * @param header
 * @return {Object} challengeMap
 */
const parseChallenges = (header) => {
    const schemeSeparator = header.indexOf(' ');
    const challenges = header.substring(schemeSeparator + 1).split(', ');
    const challengeMap = {};

    challenges.forEach((challenge) => {
        const [key, value] = challenge.split('=');
        challengeMap[key.trim()] = window.decodeURI(value.replace(/(^"|"$)/g, ''));
    });

    return challengeMap;
}