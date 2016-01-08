/* Common app functionality */

class App {
    private initialized: boolean = false;

    // Common initialization function (to be called from each page)
    initialize() {
        $('body').append(
            '<div id="notification-message">' +
            '<div class="padding">' +
            '<div id="notification-message-close"></div>' +
            '<div id="notification-message-header"></div>' +
            '<div id="notification-message-body"></div>' +
            '</div>' +
            '</div>');

        // After initialization, enable the showNotification function
        this.initialized = true;
    }

    // Notification function, enabled after initialization
    showNotification(header: string, text: string) {
        if (!this.initialized) {
            console.log('Add-in has not yet been initialized.');
            return;
        }

        $('#notification-message-header').text(header);
        $('#notification-message-body').text(text);
        $('#notification-message').slideDown('fast');
    }
}

var app = new App();