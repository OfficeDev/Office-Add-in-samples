/**
 * This method stores the claim challenge to the localStorage in the browser to be used when acquiring a token
 * @param {String} claimsChallenge
 */
const addClaimsToStorage = (claimsChallenge, claimsChallengeId) => {
    sessionStorage.setItem(claimsChallengeId, claimsChallenge);
};

/**
 * This method retrieves the claims challenge from the localStorage
 * @param {string} claimsChallengeId 
 * @returns 
 */
const getClaimsFromStorage = (claimsChallengeId) => {
    return sessionStorage.getItem(claimsChallengeId);
};

/**
 * This method clears localStorage of any claims challenge entry
 * @param {Object} account
 */
const clearStorage = (account) => {
    for (var key in sessionStorage) {
        if (key.startsWith(`cc.${msalConfig.auth.clientId}.${account.idTokenClaims.oid}`))
            sessionStorage.removeItem(key);
    }
};
