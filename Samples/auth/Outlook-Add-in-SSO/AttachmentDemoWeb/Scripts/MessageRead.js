// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

(function () {
    "use strict";

    var messageBanner;
    var overlay;
    var spinner;
    var retryGetAccessToken = 0;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric['MessageBanner'](element);

            var overlayComponent = document.querySelector(".ms-Overlay");
            // Override click so user can't dismiss overlay
            overlayComponent.addEventListener("click", function (e) {
                e.preventDefault();
                e.stopImmediatePropagation();
            });
            overlay = new window.fabric["Overlay"](overlayComponent);

            var spinnerElement = document.querySelector(".ms-Spinner");
            spinner = new window.fabric["Spinner"](spinnerElement);
            spinner.stop();

            $("#save-selected").on("click", function () {
                saveAttachmentsToOneDrive(getSelectedAttachments());
            });

            initializePane();
        });
    };

    // Initialize the pane
    function initializePane() {
        // Check if item has any attachments
        if (Office.context.mailbox.item.attachments.length > 0) {
            $("#main-content").show();
            populateList();
            initListItems();
        } else {
            $("#no-attachments").show();
        }
    }

    function populateList() {
        // Get the list
        var attachmentList = $(".ms-List");

        Office.context.mailbox.item.attachments.forEach(function (attachment) {
            var listItem = $("<li>")
                .addClass("ms-ListItem")
                .addClass("is-selectable")
                .attr("tabindex", "0")
                .appendTo(attachmentList);

            $("<div>")
                .addClass("attachment-id")
                .text(attachment.id)
                .appendTo(listItem);

            $("<span>")
                .addClass("ms-ListItem-secondaryText")
                .text(attachment.name)
                .appendTo(listItem);

            var contentType = attachment.attachmentType === "file" ?
                attachment.contentType : "Outlook item";
            if (contentType === null || contentType.length === 0) {
                contentType = "unknown";
            }

            $("<span>")
                .addClass("ms-ListItem-secondaryText")
                .text(contentType)
                .appendTo(listItem);

            $("<span>")
                .addClass("ms-ListItem-metaText")
                .text(generateFileSizeString(attachment.size))
                .appendTo(listItem);

            $("<div>")
                .addClass("ms-ListItem-selectionTarget")
                .appendTo(listItem);

            var actions = $("<div>")
                .addClass("ms-ListItem-actions")
                .appendTo(listItem);

            var saveAction = $("<div>")
                .addClass("ms-ListItem-action")
                .appendTo(actions);

            $("<i>")
                .addClass("ms-Icon")
                .addClass("ms-Icon--Save")
                .appendTo(saveAction);
        });
    }

    function generateFileSizeString(size) {
        if (size > 1048576) {
            var megString = Math.round(size / 1048576).toString() + " MB";
            return megString;
        }

        if (size > 1024) {
            var kbString = Math.round(size / 1024).toString() + " KB";
            return kbString;
        }

        else return size.toString() + " B";
    }

    function initListItems() {
        var ListElements = document.querySelectorAll(".ms-List");
        for (var i = 0; i < ListElements.length; i++) {
            new fabric['List'](ListElements[i]);
        }

        $(".ms-ListItem-selectionTarget").on("click", function () {
            var disableButton = $(".is-selected").length === 0;
            $("#save-selected").prop("disabled", disableButton);
        });

        $(".ms-ListItem-action").on("click", function () {
            var attachmentId = $(this).closest(".ms-ListItem").children(".attachment-id").text();
            saveAttachmentsToOneDrive([getRestId(attachmentId)]);
        });
    }

    function getSelectedAttachments() {
        var selectedItems = $(".is-selected");
        if (selectedItems.length > 0) {
            var attachmentIds = [];

            for (var i = 0; i < selectedItems.length; i++) {
                var id = $(selectedItems[i]).children(".attachment-id").text();
                attachmentIds.push(getRestId(id));
            }
        }
        return attachmentIds;
    }

    async function saveAttachmentsToOneDrive(attachmentIds, options) {
        //Set default SSO options if they are not provided
        if (options === undefined) options = { allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true };

        showSpinner();

        // Attempt to get an SSO token
        try {
            let bootstrapToken = await OfficeRuntime.auth.getAccessToken(options);

            // The /api/saveAttachmentsToOneDrive controller will make the token exchange and use the 
            // access token it gets back to make the call to MS Graph.
            // Server-side errors are caught in the .fail block of saveAttachmentsWithSSO.
            saveAttachmentsWithSSO("/api/saveAttachments", bootstrapToken, attachmentIds);
        }
        catch (exception) {
            // The only exceptions caught here are exceptions in your code in the try block
            // and errors returned from the call of `getAccessToken` above.
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showNotification("EXCEPTION: ", JSON.stringify(exception));
            }
        }
    }

    function saveAttachmentsWithSSO(relativeURL, accessToken, attachmentIds) {

        var saveAttachmentsRequest = {
            attachmentIds: attachmentIds,
            messageId: getRestId(Office.context.mailbox.item.itemId)
        };

        $.ajax({
            url: relativeURL,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "POST",
            data: JSON.stringify(saveAttachmentsRequest),
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            showNotification("Success", "Attachments saved");
        }).fail(function (error) {
            handleServerSideErrors(result);
        }).always(function () {
            hideSpinner();
        });
    }

    function handleClientSideErrors(error) {
        switch (error.code) {

            case 13001:
                // No one is signed into Office. If the add-in cannot be effectively used when no one 
                // is logged into Office, then the first call of getAccessToken should pass the 
                // `allowSignInPrompt: true` option.
                showNotification("Client side error 13001:", "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");
                break;
            case 13002:
                // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
                // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
                showNotification("Client side error 13002:", "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again.");
                break;
            case 13006:
                // Only seen in Office on the web.
                showNotification("Client side error 13006:", "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again.");
                break;
            case 13008:
                // Only seen in Office on the web.
                showNotification("Client side error 13008:", "Office is still working on the last operation. When it completes, try this operation again.");
                break;
            case 13010:
                // Only seen in Office on the web.
                showNotification("Client side error 13010:", "Follow the instructions to change your browser's zone configuration.");
                break;
            default:
                // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
                // to non-SSO sign-in.
                dialogFallback();
                break;
        }
    }

    function handleServerSideErrors(result) {

        // Our special handling on the server will cause the result that is returned
        // from a AADSTS50076 (a 2FA challenge) to have a Message property but no ExceptionMessage.
        var message = JSON.parse(result.responseText).Message;


        // Results from other errors (other than AADSTS50076) will have an ExceptionMessage property.
        var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;

        // Microsoft Graph requires an additional form of authentication. Have the Office host 
        // get a new token using the Claims string, which tells AAD to prompt the user for all 
        // required forms of authentication.
        if (message) {
            if (message.indexOf("AADSTS50076") !== -1) {
                var claims = JSON.parse(message).Claims;
                var claimsAsString = JSON.stringify(claims);
                saveAttachmentsToOneDrive(getSelectedAttachments(), { authChallenge: claimsAsString });
                return;
            }
        }

        if (exceptionMessage) {

            // On rare occasions the bootstrap token is unexpired when Office validates it,
            // but expires by the time it is sent to AAD for exchange. AAD will respond
            // with "The provided value for the 'assertion' is not valid. The assertion has expired."
            // Retry the call of getAccessToken (no more than once). This time Office will return a 
            // new unexpired bootstrap token.
            if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
                && (retryGetAccessToken <= 0)) {
                retryGetAccessToken++;
                saveAttachmentsToOneDrive(getSelectedAttachments());
            }
            else {
                // For debugging: 
                // showResult(["AAD ERROR: " + JSON.stringify(exchangeResponse)]);  

                // For all other AAD errors, fallback to non-SSO sign-in.                            
                dialogFallback();
            }
        }
    }


    // Helper function to show spinner
    function showSpinner() {
        spinner.start();
        overlay.show();
    }

    // Helper function to hide spinner
    function hideSpinner() {
        spinner.stop();
        overlay.hide();
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
    }


    // Dialog API

    var loginDialog;

    function dialogFallback() {

        var url = "/azureadauth/login";
        showLoginPopup(url);
    }

    // This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
    // and access token provider.
    function processMessage(arg) {

        console.log("Message received in processMessage: " + JSON.stringify(arg));
        let message = JSON.parse(arg.message);

        if (message.status === "success") {
            // We now have a valid access token.
            loginDialog.close();
            let attachmentIds = getSelectedAttachments();
            saveAttachmentsWithSSO("/api/saveAttachmentsFallback", message.accessToken, attachmentIds);
        } else {
            // Something went wrong with authentication or the authorization of the web application.
            loginDialog.close();
            showNotification("Error during fallback authorization:", "Unable to successfully authenticate user or authorize application. Error is: " + message.error);
        }
    }

    // Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
    function showLoginPopup(url) {
        var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + url;

        // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
        Office.context.ui.displayDialogAsync(fullUrl,
            { height: 60, width: 30 }, function (result) {
                console.log("Dialog has initialized. Wiring up events");
                loginDialog = result.value;
                loginDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
            });
    }


})();

