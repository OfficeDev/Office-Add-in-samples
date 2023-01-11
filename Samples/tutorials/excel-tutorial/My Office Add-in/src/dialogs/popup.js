(function () {
    "use strict";

    Office.onReady()
        .then(function () {
            document.getElementById("ok-button").onclick = sendStringToParentPage;
        });

    function sendStringToParentPage() {
        const userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
}());