var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.
 *
 * This file shows how to use the SSO API to get a bootstrap token.
 */
// If the add-in is running in Internet Explorer, the code must add support 
// for Promises.
if (!window.Promise) {
    window.Promise = Office.Promise;
}
Office.onReady(function (info) {
    $(document).ready(function () {
        $('#getGraphDataButton').click(getGraphData);
    });
});
var retryGetAccessToken = 0;
function getGraphData() {
    return __awaiter(this, void 0, void 0, function () {
        var bootstrapToken, exchangeResponse, mfaBootstrapToken, exception_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 6, , 7]);
                    return [4 /*yield*/, OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true })];
                case 1:
                    bootstrapToken = _a.sent();
                    return [4 /*yield*/, getGraphToken(bootstrapToken)];
                case 2:
                    exchangeResponse = _a.sent();
                    if (!exchangeResponse.claims) return [3 /*break*/, 5];
                    return [4 /*yield*/, OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims })];
                case 3:
                    mfaBootstrapToken = _a.sent();
                    return [4 /*yield*/, getGraphToken(mfaBootstrapToken)];
                case 4:
                    exchangeResponse = _a.sent();
                    _a.label = 5;
                case 5:
                    if (exchangeResponse.error) {
                        // AAD errors are returned to the client with HTTP code 200, so they do not trigger
                        // the catch block below.
                        handleAADErrors(exchangeResponse);
                    }
                    else {
                        // For debugging:
                        // showMessage("ACCESS TOKEN: " + JSON.stringify(exchangeResponse.access_token));
                        // makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
                        // in the .fail callback of that call, not in the catch block below.
                        makeGraphApiCall(exchangeResponse.access_token);
                    }
                    return [3 /*break*/, 7];
                case 6:
                    exception_1 = _a.sent();
                    // The only exceptions caught here are exceptions in your code in the try block
                    // and errors returned from the call of `getAccessToken` above.
                    if (exception_1.code) {
                        handleClientSideErrors(exception_1);
                    }
                    else {
                        showMessage("EXCEPTION: " + JSON.stringify(exception_1));
                    }
                    return [3 /*break*/, 7];
                case 7: return [2 /*return*/];
            }
        });
    });
}
function getGraphToken(bootstrapToken) {
    return __awaiter(this, void 0, void 0, function () {
        var response;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, $.ajax({ type: "GET",
                        url: "/auth",
                        headers: { "Authorization": "Bearer " + bootstrapToken },
                        cache: false
                    })];
                case 1:
                    response = _a.sent();
                    return [2 /*return*/, response];
            }
        });
    });
}
function handleClientSideErrors(error) {
    switch (error.code) {
        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one 
            // is logged into Office, then the first call of getAccessToken should pass the 
            // `allowSignInPrompt: true` option. Since this sample does that, you should not see this error
            showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.");
            break;
        case 13006:
            // Only seen in Office on the web.
            showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again.");
            break;
        case 13008:
            // Only seen in Office on the web.
            showMessage("Office is still working on the last operation. When it completes, try this operation again.");
            break;
        case 13010:
            // Only seen in Office on the web.
            showMessage("Follow the instructions to change your browser's zone configuration.");
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in.
            dialogFallback();
            break;
    }
}
function handleAADErrors(exchangeResponse) {
    // On rare occasions the bootstrap token is unexpired when Office validates it,
    // but expires by the time it is sent to AAD for exchange. AAD will respond
    // with "The provided value for the 'assertion' is not valid. The assertion has expired."
    // Retry the call of getAccessToken (no more than once). This time Office will return a 
    // new unexpired bootstrap token. 
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
            (retryGetAccessToken <= 0)) {
        retryGetAccessToken++;
        getGraphData();
    }
    else {
        // For all other AAD errors, fallback to non-SSO sign-in.
        // For debugging: 
        // showMessage("AAD ERROR: " + JSON.stringify(exchangeResponse));                   
        dialogFallback();
    }
}
//# sourceMappingURL=ssoAuth.js.map