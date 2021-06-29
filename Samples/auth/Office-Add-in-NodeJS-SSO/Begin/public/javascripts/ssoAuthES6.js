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

Office.onReady(function(info) {
    $(function() {
        $('#getGraphDataButton').on('click',getGraphData);
    });
});

