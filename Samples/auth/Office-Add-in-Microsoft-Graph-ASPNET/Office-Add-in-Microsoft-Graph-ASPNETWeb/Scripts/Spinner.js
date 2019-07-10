// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

"use strict";

Office.initialize = function () {

    $(document).ready(function () {

        app.initialize();

        // Office UI Fabric spinner initialization code.
        if (typeof fabric === "object") {
            if ('Spinner' in fabric) {
                var element = document.querySelector('.ms-Spinner');
                new fabric['Spinner'](element);
            }
        }
    });
};
