(function () {
    "use strict";

    Office.onReady()
        .then(function () {
            document.getElementById("ok-button").onclick = sendStringToParentPage;
        });

    /**
     * This sends the text input from the dialog back to the parent.
     */
    function sendStringToParentPage() {
        const userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
}());