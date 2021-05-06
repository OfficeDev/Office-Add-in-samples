// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function getRestUrl() {
    if (Office.context.mailbox.restUrl !== undefined) {
        return Office.context.mailbox.restUrl;
    } else {
        // Just assume Office 365
        return "https://outlook.office.com/api"
    }
}

// Helper function to convert EWS ID format
// to REST format
function getRestId(ewsId) {
    return Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
}