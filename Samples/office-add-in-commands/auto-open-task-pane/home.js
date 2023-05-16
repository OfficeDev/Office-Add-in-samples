Office.onReady((reason) => {

    document.getElementById('template-description').innerText =
        'This sample show how to programmatically tag a document to auto-open a task pane.';

    // Add click event handlers for the buttons.
    document.getElementById('set-auto-open-on').onclick = setAutoOpenOn;
    document.getElementById('set-auto-open-off').onclick = setAutoOpenOff;
    document.getElementById('turn-off-message').onclick = turnOffMessage;

    turnOffMessage(); //Ensure message footer is not showing on start.
});

function setAutoOpenOn() {
    Office.context.document.settings.set(
        'Office.AutoShowTaskpaneWithDocument',
        true
    );
    Office.context.document.settings.saveAsync();
    showNotification(
        'The auto-open setting has been set to ON for this document.'
    );
}

function setAutoOpenOff() {
    Office.context.document.settings.remove(
        'Office.AutoShowTaskpaneWithDocument'
    );
    Office.context.document.settings.saveAsync();
    showNotification(
        'The auto-open setting has been set to OFF for this document.'
    );
}

// Helper function for displaying notifications in the footer.
function showNotification(content) {
    document.getElementById('message-text').innerText = content;
    document.getElementById('message-text').style.visibility = 'visible';
    document.getElementById('turn-off-message').style.visibility = 'visible';
}

/**
 * Hides the message footer.
 */
function turnOffMessage() {
    document.getElementById('message-text').style.visibility = 'hidden';
    document.getElementById('turn-off-message').style.visibility = 'hidden';
}
