(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
        });
    };
})();

function showUnicode() {
    Excel.run(function (ctx) {
        const range = ctx.workbook.getSelectedRange();
        range.load("values");
        return ctx.sync(range).then(function (range) {
            const url = "https://localhost:44360/api/analyzeunicode?value=" + range.values[0][0];
            $.ajax({
                type: "GET",
                url: url,
                success: function (data) {
                    let htmlData = data.replace(/\r\n/g, '<br>');
                    $("#txtResult").html(htmlData);
                },
                error: function (data) {
                    $("#txtResult").html("error occurred in ajax call.");
                }
            });
        });
    });
}