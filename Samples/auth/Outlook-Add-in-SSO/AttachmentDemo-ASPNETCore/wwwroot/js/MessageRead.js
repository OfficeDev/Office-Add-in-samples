// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

(function () {
    "use strict";

    // Set authSSO to false to force using the MSAL fallback path.
    let authSSO = true;

    Office.onReady(() => {
        $(document).ready(() => {
            $("#save-button").on("click", function () {
                saveAttachments();
            });
            initializePane();
        });
    });

    // Helper function to convert EWS ID format
    // to REST format
    function getRestId(ewsId) {
        return Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
    }

    /**
     * Initialize the task pane.
     */
    function initializePane() {
        // List any attachments in the task pane.
        if (Office.context.mailbox.item.attachments.length > 0) {
            $("#save-button").show();
            Office.context.mailbox.item.attachments.forEach(function (attachment) {
                $("#list-of-attachments").append(
                    "<fluent-checkbox class='attachment-item' value='" +
                    attachment.id +
                    "'>" +
                    attachment.name +
                    "</fluent-checkbox>"
                );
            });
            $("#list-of-attachments").show();
        } else {
            $("#no-attachments").show();
        }
    }

    /**
     *  For each attachment selected by the user, it is added to an
     *   array that is returned to the caller.
     * @returns An array of checked items.
     */
    function getSelectedAttachments() {
        let checkedItems = [];
        $(".attachment-item").each((index, item) => {
            checkedItems.push(getRestId(item.value));
        });
        return checkedItems;
    }

    /**
     * Manages the process to save attachments.
     */
    async function saveAttachments() {
        clearMessage(); // Clear message log on task pane each time an API runs.

        //Check that there is a selection.
        const selectedAttachments = getSelectedAttachments();
        if (selectedAttachments.length === 0) {
            showMessage("Please select one or more attachments to save.");
            return;
        }

        $("#spinner").show();
        let saveAttachmentsRequest = {
            attachmentIds: selectedAttachments,
            messageId: getRestId(Office.context.mailbox.item.itemId),
        };

        const result = await postRestApi(
            "/api/saveAttachments",
            JSON.stringify(saveAttachmentsRequest)
        );
        if (result !== null) showMessage("Success: Attachments saved");
        $("#spinner").hide();
        return;
    }

    /**
     * Handles errors from getAccessToken (SSO).
     * For most errors display a message with information for the user.
     * @param {*} error The error from getAccessToken. 
     * @returns true if the caller should switch to MSAL fallback for sign in; otherwise, false.
     */
    function handleClientSideErrors(error) {
        switch (error.code) {
            case 13001:
                // No one is signed into Office. If the add-in cannot be effectively used when no one
                // is logged into Office, then the first call of getAccessToken should pass the
                // `allowSignInPrompt: true` option.
                showMessage(
                    "Client side error 13001: No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
                );
                break;
            case 13002:
                // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
                // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
                showMessage(
                    "Client side error 13002: You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
                );
                break;
            case 13006:
                // Only seen in Office on the web.
                showMessage(
                    "Client side error 13006: Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
                );
                break;
            case 13008:
                // Only seen in Office on the web.
                showMessage(
                    "Client side error 13008: Office is still working on the last operation. When it completes, try this operation again."
                );
                break;
            case 13010:
                // Only seen in Office on the web.
                showMessage(
                    "Client side error 13010: Follow the instructions to change your browser's zone configuration."
                );
                break;
            default:
                // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back to MSAL sign-in.
                return true;
                showMessage(
                    "Client side error: " +
                    error.code +
                    error.message +
                    ": You would need to fall back to MSAL sign-in to handle this error."
                );
                break;
        }
        return false;
    }

    /**
     * Handles errors returned from REST APIs on the server.
     * @param {*} error The error returned from the REST API call.
     * @returns true if getAccessToken should be called to get a refreshed token; otherwise, false.
     */
    async function handleServerSideErrors(error) {
        try {
            // On rare occasions the access token is unexpired when Office validates it,
            // but expires by the time it is sent to Microsoft identity for the OBO flow.
            // Microsoft identity returns an error which will be returned to this task pane.
            // Retry the call of getAccessToken (no more than once). This time Office will return a
            // new unexpired access token.
            const authHeader = await error.getResponseHeader("www-authenticate");
            if (
                authHeader !== null &&
                authHeader.includes("invalid_token") &&
                authHeader.includes("The token expired")
            ) {
                return true; // Caller should get refreshed token from getAccessToken.
            } else {
                // Handle unexpected www-authenticate header.
                if (authHeader !== null) {
                    showMessage("There was an unexpected error. Details from response header: " + authHeader);
                    return false;
                }
                // Otherwise proceed to next error check.
            }

            // This section handles errors that were handled by the controller and returned as JSON.
            const message = error.responseJSON.value.message;
            const details = error.responseJSON.value.details;
            const code = error.responseJSON.value.code;
            showMessage(
                "Server side error: " +
                message +
                "<br/>Details: " +
                JSON.stringify(details)
            );
            if (code !== undefined) showMessage("Error code: " + code);
        } catch (e) {
            // For any errors we couldn't handle, report an error.
            showMessage("There was an unexpected server side error.");
        }
        return false;
    }

    /**
     * Makes the AJAX call to the requested REST API.
     * jsonData must be a string in JSON format.
     * @param {*} urlData The URL to call.
     * @param {*} jsonData The JSON data to pass to the REST API.
     * @param {*} retryRequest true if this is a second recursive call to get a refreshed access token; otherwise, false.
     * @returns 
     */
    async function postRestApi(urlData, jsonData, retryRequest) {
        if (retryRequest === undefined) retryRequest = false;
        try {
            let isConsent = false;
            const accessToken = await getAccessToken(retryRequest);
            if (accessToken === null) return null;
            const result = await $.ajax({
                type: "POST",
                contentType: "application/json; charset=utf-8",
                url: urlData,
                headers: { Authorization: "Bearer " + accessToken },
                cache: false,
                data: jsonData,
                success: function (result, status, request) {
                    // Check for an incremental consent request (which is returned as 200 success).
                    const header = request.getResponseHeader("www-authenticate");
                    if (header !== null) {
                        showMessage("Incremental consent may be required. Check that your app registration has the correct permissions for Microsoft Graph and Microsoft identity. Details from MSAL response header: " + header);
                        isConsent = true;
                        // For more information about incremental consent, see https://github.com/AzureAD/microsoft-identity-web/wiki/Managing-incremental-consent-and-conditional-access
                    }
                },
            });
            if (isConsent) return null;
            return result;
        } catch (error) {
            // In the scenario where the token expired we will retry once to get a refreshed token.
            const retry = await handleServerSideErrors(error);
            if (retry && !retryRequest) {
                postRestApi(urlData, jsonData, true);
            }
            return null;
        }
    }

    /**
     * Returns the access token to the REST API server. If getAccessToken fails, falls back to use MSAL sign in.
     * Returns null if an error was handled and the user has information they need to follow up on.
     * Throws an error if sign-in fails.
     * 
     * @returns An access token if successful; otherwise, null.
     */
    async function getAccessToken() {
        if (authSSO) {
            try {
                // Attempt to get token through SSO.
                const accessToken = await Office.auth.getAccessToken({
                    allowSignInPrompt: true,
                    allowConsentPrompt: true,
                    forMSGraphAccess: true,
                });
                return accessToken;
            } catch (error) {
                let useFallback = handleClientSideErrors(error);
                if (useFallback) {
                    // Set authSSO flag to false to force fallback to MSAL later in this function.
                    authSSO = false;
                } else {
                    // If handleClientSideErrors returned false, then instructions were posted for the user to follow.
                    // In this case return null.
                    return null;
                }
            }
        }
        if (!authSSO) {
            // Attempt to get token through MSAL fallback.
            try {
                const accessToken = await getAccessTokenMSAL();
                return accessToken;
            } catch (error) {
                // if both SSO and fallback fail, throw the error to the caller.
                throw error;
            }
        }
    }

    /**
     * Gets an access token to the REST API server by using MSAL (Microsoft Authentication Library for the browser.)
     * @returns A promise which if successful, returns the access token.
     */
    async function getAccessTokenMSAL() {
        // Create a promise to wrap the dialog callback we need to process later in this function.
        let promise = await new Promise((resolve, reject) => {
            const url = "/dialog.html";
            var fullUrl =
                location.protocol +
                "//" +
                location.hostname +
                (location.port ? ":" + location.port : "") +
                url;

            // height and width are percentages of the size of the parent Office application, e.g., Outlook, PowerPoint, Excel, Word, etc.
            Office.context.ui.displayDialogAsync(
                fullUrl,
                { height: 60, width: 30 },
                function (result) {
                    console.log("Dialog has initialized. Wiring up events");
                    let loginDialog = result.value;
                    loginDialog.addEventHandler(
                        Office.EventType.DialogMessageReceived,
                        function processMessage2(arg) {
                            console.log("Message received in processMessage");
                            let messageFromDialog = JSON.parse(arg.message);

                            if (messageFromDialog.status === "success") {
                                // We now have a valid access token.
                                loginDialog.close();

                                // Add access token to the client request and run the callback function.
                                resolve(messageFromDialog.result);
                            } else {
                                // Something went wrong with authentication or the authorization of the web application.
                                loginDialog.close();
                                reject(JSON.stringify(error.toString()));
                            }
                        }
                    );
                }
            );
        });
        return promise;
    }
})();
