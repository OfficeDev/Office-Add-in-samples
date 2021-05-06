// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
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