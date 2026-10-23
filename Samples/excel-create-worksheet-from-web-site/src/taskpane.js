/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('set-auto-open-on').onclick = setAutoOpenOn;
        document.getElementById('set-auto-open-off').onclick = setAutoOpenOff;
        document.getElementById('turn-off-message').onclick = turnOffMessage;

        turnOffMessage(); // Ensure message footer isn't showing on start.
    }
});

function setAutoOpenOn() {
    Office.context.document.settings.set(
        'Office.AutoShowTaskpaneWithDocument',
        true
    );
    Office.context.document.settings.saveAsync();
    showNotification(
        'The auto-open setting has been set to ON for this workbook.'
    );
}

function setAutoOpenOff() {
    Office.context.document.settings.remove(
        'Office.AutoShowTaskpaneWithDocument'
    );
    Office.context.document.settings.saveAsync();
    showNotification(
        'The auto-open setting has been set to OFF for this workbook.'
    );
}

// Helper function for displaying notifications in the footer.
function showNotification(content) {
    const footer = document.querySelector('footer');
    const messageText = document.getElementById('message-text');

    
    messageText.innerText = content;
    footer.style.display = 'block';
}

/**
 * Hides the message footer.
 */
function turnOffMessage() {
    const footer = document.querySelector('footer');
    footer.style.display = 'none';
}
