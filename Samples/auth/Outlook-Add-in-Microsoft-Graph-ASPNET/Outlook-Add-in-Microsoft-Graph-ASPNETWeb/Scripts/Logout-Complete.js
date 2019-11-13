// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
The pop-up dialog uses this method to tell the task pane that the user is logged out.
*/

Office.initialize = function (reason) {
    $(document).ready(function () {
        Office.context.ui.messageParent("success");
    });
}