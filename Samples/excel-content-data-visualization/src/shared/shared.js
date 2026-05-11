let shared = (function () {
    "use strict";

    let shared = {};
    shared.bindingID = "myBinding";

    // Common initialization function (to be called from each page).
    shared.initialize = function () {
        // Prevent duplicate initialization.
        if (document.getElementById('notification-message')) {
            return;
        }

        document.body.insertAdjacentHTML('beforeend',
            '<div id="notification-message">' +
            '<div class="padding">' +
            '<div id="notification-message-close"></div>' +
            '<div id="notification-message-header"></div>' +
            '<div id="notification-message-body"></div>' +
            '</div>' +
            '</div>');

        document.getElementById('notification-message-close').addEventListener('click', function () {
            document.getElementById('notification-message').style.display = 'none';
        });


        // After initialization, expose a common notification function.
        shared.showNotification = function (header, text) {
            document.getElementById('notification-message-header').textContent = header;
            document.getElementById('notification-message-body').textContent = text;
            document.getElementById('notification-message').style.display = 'block';
        };
    };

    return shared;
})();