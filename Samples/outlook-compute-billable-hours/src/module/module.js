/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(function () {
    "use strict"
    // The initialze function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            setBillingRates(127);
        });
    };
})();

// The next three functions change the billing rate used to
// calculate the totals for each item.
function changeToAssociateRate(event) {
    setBillingRates(127);
    completeEvent(event);
};

function changeToPartnerRate(event) {
    setBillingRates(237);
    completeEvent(event)
};

function changeToExecutiveRate(event) {
    setBillingRates(449);
    completeEvent(event);
};

function completeEvent(event) {
    if (event) {
        event.completed(true);
    }
};

var runningTotalHours;
var runningTotalAmount;

// Updates the user interface with the calculated
// amounts to bill for each item.
function setBillingRates(rate) {
    runningTotalHours = 0;
    runningTotalAmount = 0;
    var xhr = [
        setMeetingsRate(rate),
        setEmailRate(rate),
        setTasksRate(rate)
    ];

    $.when(xhr[0], xhr[1], xhr[2]).then(function () {
        showGrandTotal();
    });
};

// Constructs the meetings table and calculated the total
// billing amount for each item and for all meetings.
function setMeetingsRate(rate) {
    return $.getJSON("../../assets/sampleMeetingData.json", function (data) {
        var jsonData = data.Appointments;

        var dataTable = $("#app-Meetings-table");
        dataTable.html("");

        var headerRow = $('<div />');
        headerRow.append(makeHeaderCell("Subject", "5"));
        headerRow.append(makeHeaderCell("Attendees", "5"));
        headerRow.append(makeHeaderCell("Hours", "1", "true"));
        headerRow.append(makeHeaderCell("Total", "1", "true"));

        dataTable.append(headerRow);

        var totalHours = 0;
        var totalAmount = 0;

        for (var i in jsonData) {
            var dataRow = $("<div />", {
                "class": "ms-Grid-row app-Grid-row"
            });
            dataRow.append(makeRowCell(jsonData[i].Subject, "5"));
            dataRow.append(makeRowCell(jsonData[i].Attendees, "5"));
            dataRow.append(makeRowCell(jsonData[i].Hours, "1", "true"));
            dataRow.append(makeRowCell(jsonData[i].Hours * rate, "1", "true"));

            totalHours += Number(jsonData[i].Hours);
            totalAmount += rate * (jsonData[i].Hours);

            dataTable.append(dataRow);
        }
        
        dataTable.append(makeTotalRow(totalHours, totalAmount));

        runningTotalHours += totalHours;
        runningTotalAmount += totalAmount;
    });
};

// Constructs the email table and calculated the total
// billing amount for each item and for all email.
function setEmailRate(rate) {
    return $.getJSON("../../assets/sampleEmailData.json", function (data) {
        var jsonData = data.Messages;

        var dataTable = $("#app-Email-table");
        dataTable.html("");

        var headerRow = $('<div />');
        headerRow.append(makeHeaderCell("Subject", "5"));
        headerRow.append(makeHeaderCell("Recipients", "5"));
        headerRow.append(makeHeaderCell("Hours", "1", "true"));
        headerRow.append(makeHeaderCell("Total", "1", "true"));

        dataTable.append(headerRow);

        var totalHours = 0;
        var totalAmount = 0;

        for (var i in jsonData) {
            var dataRow = $("<div />", {
                "class": "ms-Grid-row app-Grid-row"
            });

            dataRow.append(makeRowCell(jsonData[i].Subject, "5"));
            dataRow.append(makeRowCell(jsonData[i].Recipients, "5"));
            dataRow.append(makeRowCell(jsonData[i].Hours, "1", "true"));
            dataRow.append(makeRowCell(jsonData[i].Hours * rate, "1", "true"));

            totalHours += Number(jsonData[i].Hours);
            totalAmount += rate * (jsonData[i].Hours);

            dataTable.append(dataRow);
        }

        dataTable.append(makeTotalRow(totalHours, totalAmount));

        runningTotalHours += totalHours;
        runningTotalAmount += totalAmount;
    });
};

// Constructs the tasks table and calculated the total
// billing amount for each item and for all tasks.
function setTasksRate(rate) {
    return $.getJSON("../../assets/sampleTaskData.json", function (data) {
        var jsonData = data.Tasks;

        var dataTable = $("#app-Tasks-table");
        dataTable.html("");

        var headerRow = $('<div />');
        headerRow.append(makeHeaderCell("Action", "10"));
        headerRow.append(makeHeaderCell("Hours", "1", "true"));
        headerRow.append(makeHeaderCell("Total", "1", "true"));

        dataTable.append(headerRow);

        var totalHours = 0;
        var totalAmount = 0;

        for (var i in jsonData) {
            var dataRow = $("<div />", {
                "class": "ms-Grid-row app-Grid-row"
            });
            dataRow.append(makeRowCell(jsonData[i].Action, "10"));
            dataRow.append(makeRowCell(jsonData[i].Hours, "1", "true"));
            dataRow.append(makeRowCell(jsonData[i].Hours * rate, "1", "true"));

            totalHours += Number(jsonData[i].Hours);
            totalAmount += rate * (jsonData[i].Hours);

            dataTable.append(dataRow);
        }

        dataTable.append(makeTotalRow(totalHours, totalAmount));

        runningTotalHours += totalHours;
        runningTotalAmount += totalAmount;
    });
};

// Creates the HTML for displaying a table header cell.
function makeHeaderCell(text, width, right) {
    var cssClass = "ms-Grid-col ms-fontColor-themeDark ms-font-l ms-u-lg" + width;

    if (right) {
        cssClass += " app-Cell-right";
    }

    return $("<div />", {
        "class": cssClass,
        "html": text
    })
};

// Creates the HTML for displaying a table cell.
function makeRowCell(text, width, right) {
    var cssClass = "ms-Grid-col ms-u-md4 ms-u-lg" + width;

    if (right) {
        cssClass += " app-Cell-right";
    }

	return $("<div />", {
        "class" : cssClass,
        "html"  : text
    });    
}

function makeTotalRow(totalHours, totalAmount) {
        var totalRow = $("<div />", {
            "class": "ms-Grid-row  ms-fontColor-themeDark ms-font-l app-Grid-row"
        });

        totalRow.append($("<div />", {
            "class": "app-Cell-right ms-Grid-col ms-u-lg10",
            "html": "Totals:"
        }));

        totalRow.append($("<div />", {
            "class": "app-Cell-right ms-Grid-col ms-u-lg1",
            "html": totalHours
        }));

        totalRow.append($("<div />", {
            "class": "app-Cell-right ms-Grid-col ms-u-lg1",
            "html": totalAmount
        }));

    return totalRow;
}

// Creates the row that displays the grand total for the page.
function showGrandTotal() {
    var totalTable = $("#app-Running-total");
    totalTable.html("");

     var totalRow = $("<div />", {
         "class": "app-Title-bar ms-bgColor-themeDarker ms-fontColor-themeLighter ms-font-xxl ms-fontWeight-semibold"
     });

     totalRow.append($("<div />", {
         "class": "app-Cell-right ms-Grid-col ms-u-lg10",
         "html": "Grand total:"
     }));

     totalRow.append($("<div />", {
         "class": "app-Cell-right ms-Grid-col ms-u-lg1",
         "html": runningTotalHours
     }));

     totalRow.append($("<div />", {
         "class": "app-Cell-right ms-Grid-col ms-u-lg1",
         "html": runningTotalAmount
     }));

    totalTable.append(totalRow);
};