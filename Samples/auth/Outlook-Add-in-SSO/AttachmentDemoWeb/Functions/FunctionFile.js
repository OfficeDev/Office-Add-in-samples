// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
Office.initialize = function () {

}

async function saveAllAttachments(event) {
    var attachmentIds = [];
    let accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });

    Office.context.mailbox.item.attachments.forEach(function (attachment) {
        attachmentIds.push(getRestId(attachment.id));
    });

    var saveAttachmentsRequest = {
        attachmentIds: attachmentIds,
        messageId: getRestId(Office.context.mailbox.item.itemId)
    };
    $.ajax({
        type: "POST",
        url: "/api/SaveAttachments",
        headers: { "Authorization": "Bearer " + accessToken },
        data: JSON.stringify(saveAttachmentsRequest),
        contentType: "application/json; charset=utf-8"
    }).done(function (data) {
        showSuccess("Attachments saved");
    }).fail(function (result) {
        showError("Error saving attachments");
    }).always(function () {
        event.completed();
    });

}

function showError(message) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "errorMessage",
        message: message
    });
}

function showSuccess(message) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        message: message,
        icon: "icon16",
        persistent: false
    });
}