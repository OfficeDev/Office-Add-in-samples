/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            $('.ms-Dropdown').Dropdown();
        });


        $('#applicant-name').on('change', function (e) {
            var optionSelected = $("option:selected", this);
            var valueSelected = this.value;
            fillApplicantRelatedFields(valueSelected);
        });

        $('#policy-name').on('change', function (e) {
            var optionSelected = $("option:selected", this);
            var valueSelected = this.value;
            fillPolicyRelatedFields(valueSelected);

        });

        $('#insurance-amount').on('change', function (e) {
            var insuranceAmount = $('#insurance-amount').val();
            var sampleRate = $('#sample-rate').val();
            $('#monthly-payment').text(insuranceAmount*sampleRate/10000 + '$');
        });

        $('#save-prospect').click(saveProspect);

        createMyPropectsTrackerSheet();
        fillDropDownMenus();

        //importSampleData();
    };

    // Create the My ProspectsTracker sheet 
    function createMyPropectsTrackerSheet() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            var prospectsSheet = ctx.workbook.worksheets.getActiveWorksheet();
            prospectsSheet.name = "Agent Workspace";

            // Create strings to store all static content to display in the Prospects Tracker sheet
            var sheetTitle = "Humongous Insurance Agent Center";
            var sheetHeading1 = "In Agent Center, you can easily track and manage prospects.";

            // Add all static content to the Welcome sheet and format the text
            addContentToWorksheet(prospectsSheet, "B1:X1", sheetTitle, "SheetTitle");
            addContentToWorksheet(prospectsSheet, "B3:K3", sheetHeading1, "SheetHeading");

            //Queue commands to autofit rows and columns in the sheet
            prospectsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            prospectsSheet.getUsedRange().getEntireRow().format.autofitRows();

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Import sample data into tables in the workbook
    function importSampleData() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue commands to add a new worksheets to store all the sample data
            var agentsSheet = ctx.workbook.worksheets.add("Agents");

            // Create strings to store all static content to display in the Agents sheet
            var sheetTitle = "Humongous Insurance Agent Center";
            var sheetHeading1 = "Agents - Master List";

            //Queue a command to remove gridlines from view
            agentsSheet.getRange().format.fill.color = "white";

            // Add all static content to the Transactions sheet and format the text
            addContentToWorksheet(agentsSheet, "B1:K1", sheetTitle, "SheetTitle");
            addContentToWorksheet(agentsSheet, "B3:K3", sheetHeading1, "SheetHeading");

            // Queue a command to add a new table
            var agentsTable = ctx.workbook.tables.add('Agents!B6:B6', true);
            agentsTable.name = "AgentsTable";

            // Queue a command to set the header row
            agentsTable.getHeaderRowRange().values = [["AgentName"]];
            var tableRows = agentsTable.rows;

            tableRows.add(null, [["Aanandini Kidambi"]]);
            tableRows.add(null, [["Jordan Hopkins"]]);
            tableRows.add(null, [["Amelie Laffer"]]);
            tableRows.add(null, [["Ya-ting Lo"]]);
            tableRows.add(null, [["Chelsea Leigh"]]);
            tableRows.add(null, [["Badanika Atluri"]]);

            // Quere commands to format the table
            addContentToWorksheet(agentsSheet, "B6:B6", "", "TableHeaderRow");
            addContentToWorksheet(agentsSheet, "B7:B12", "", "TableDataRows");

            // Queue commands to auto-fit columns and rows
            agentsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            agentsSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Applicants sheet

            var applicantsSheet = ctx.workbook.worksheets.add("Applicants");

            // Create strings to store all static content to display in the Applicants sheet
            var sheetTitle = "Humongous Insurance Agent Center";
            var sheetHeading1 = "Applicants - Master List";

            //Queue a command to remove gridlines from view
            applicantsSheet.getRange().format.fill.color = "white";

            // Add all static content to the Applicants sheet and format the text
            addContentToWorksheet(applicantsSheet, "B1:K1", sheetTitle, "SheetTitle");
            addContentToWorksheet(applicantsSheet, "B3:K3", sheetHeading1, "SheetHeading");

            // Queue a command to add a new table
            var applicantsTable = ctx.workbook.tables.add('Applicants!B6:D6', true);
            applicantsTable.name = "ApplicantsTable";

            // Queue a command to set the header row
            applicantsTable.getHeaderRowRange().values = [["Applicant", "Age", "Gender"]];
            var tableRows = applicantsTable.rows;

            tableRows.add(null, [["You Chioh", "55", "Male"]]);
            tableRows.add(null, [["Roelf de Boer", "43", "Male"]]);
            tableRows.add(null, [["Isa Nuijten", "28", "Female"]]);
            tableRows.add(null, [["Hanne Clausen", "33", "Female"]]);
            tableRows.add(null, [["Amalie Frederiksen", "29", "Female"]]);
            tableRows.add(null, [["Darrell Jaime", "54", "Male"]]);
            tableRows.add(null, [["Vandana Dutta", "41", "Female"]]);
            tableRows.add(null, [["William Lyons", "22", "Male"]]);
            tableRows.add(null, [["Mara Michael", "63", "Female"]]);
            tableRows.add(null, [["Clinton Slaton", "42", "Male"]]);
            tableRows.add(null, [["Bridgett Vega", "27", "Female"]]);
            tableRows.add(null, [["Paul Oswalt", "25", "Male"]]);

            // Format the table header and data rows
            addContentToWorksheet(applicantsSheet, "B6:D6", "", "TableHeaderRow");
            addContentToWorksheet(applicantsSheet, "B7:D18", "", "TableDataRows");


            // Queue commands to auto-fit columns and rows
            applicantsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            applicantsSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Polcies sheet

            var policiesSheet = ctx.workbook.worksheets.add("Policies");

            // Create strings to store all static content to display in the Policies sheet
            var sheetTitle = "Humongous Insurance Agent Center";
            var sheetHeading1 = "Policies - Master List";

            //Queue a command to remove gridlines from view
            policiesSheet.getRange().format.fill.color = "white";

            // Add all static content to the Policies sheet and format the text
            addContentToWorksheet(policiesSheet, "B1:K1", sheetTitle, "SheetTitle");
            addContentToWorksheet(policiesSheet, "B3:K3", sheetHeading1, "SheetHeading");

            // Queue a command to add a new table
            var policiesTable = ctx.workbook.tables.add('Policies!B6:D6', true);
            policiesTable.name = "PoliciesTable";

            // Queue a command to set the header row 
            policiesTable.getHeaderRowRange().values = [["PolicyName", "Medical Exam Required", "Sample Rate For $10000"]];
            var tableRows = policiesTable.rows;

            tableRows.add(null, [["Whole Life", "No", "$42.5"]]);
            tableRows.add(null, [["Universal Life", "No", "$35.5"]]);
            tableRows.add(null, [["Variable Life", "Yes", "$28.25"]]);
            tableRows.add(null, [["Term Life", "Yes", "$4.75"]]);

            // Format the table header and data rows
            addContentToWorksheet(policiesSheet, "B6:D6", "", "TableHeaderRow");
            addContentToWorksheet(policiesSheet, "B7:D10", "", "TableDataRows");

            // Queue commands to auto-fit columns and rows
            policiesSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            policiesSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Add a new sheet to save prospects
            var savedProspectsSheet = ctx.workbook.worksheets.add("Prospects");

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }


    function fillDropDownMenus() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the transactions table
            var agentsTable = ctx.workbook.tables.getItem("AgentsTable");
            var agentNameColumn = agentsTable.columns.getItem("AgentName").getDataBodyRange().load("values");

            // Queue a command to get the transactions table
            var applicantsTable = ctx.workbook.tables.getItem("ApplicantsTable");
            var applicantNameColumn = applicantsTable.columns.getItem("Applicant").getDataBodyRange().load("values");
            var applicantAgeColumn = applicantsTable.columns.getItem("Age").getDataBodyRange().load("values");
            var applicantGenderColumn = applicantsTable.columns.getItem("Gender").getDataBodyRange().load("values");

            // Queue a command to get the transactions table
            var policiesTable = ctx.workbook.tables.getItem("PoliciesTable");
            var policyNameColumn = policiesTable.columns.getItem("Policy").getDataBodyRange().load("values");
            var policySampleRateColumn = policiesTable.columns.getItem("SampleRateFor$10000").getDataBodyRange().load("values");


            // Run all of the above queued-up commands, and return a promise to indicate task completion
            return ctx.sync().then(function () {

                var agentNames = agentNameColumn.values;

                // For each agent name, add a dropdown item in the UI
                for (var i = 0; i < agentNames.length; i++) {
                    // Create New Option.
                    var name = agentNames[i];
                    var newOption = $('<option>');
                    newOption.attr('value', name).text(name);
                    $("#agent-name").append(newOption);
                }

                var applicantNames = applicantNameColumn.values;

                // For each agent name, add a dropdown item in the UI
                for (var i = 0; i < applicantNames.length; i++) {
                    // Create New Option.
                    var name = applicantNames[i];
                    var newOption = $('<option>');
                    newOption.attr('value', name).text(name);
                    $("#applicant-name").append(newOption);
                }

                var policyNames = policyNameColumn.values;

                // For each agent name, add a dropdown item in the UI
                for (var i = 0; i < policyNames.length; i++) {
                    // Create a new option
                    var name = policyNames[i];
                    var newOption = $('<option>');
                    newOption.attr('value', name).text(name);
                    $("#policy-name").append(newOption);
                }

            });
        })
            .then(function () {
                var dropdowns = $('.ms-Dropdown');
                dropdowns.Dropdown();
                dropdowns.each(function () {
                    var titles = $(this).find('.ms-Dropdown-title');
                    var items = $(this).find('ms-Dropdown-title');
                    $(titles.splice(0, titles.length - 1)).remove();
                    $(items.splice(0, items.length - 1)).remove();
                });
            })
            .catch(function (error) {

                handleError(error);
            });
    }


    function fillApplicantRelatedFields(selectedApplicant) {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the transactions table
            var applicantsTable = ctx.workbook.tables.getItem("ApplicantsTable");
            var applicantNameColumn = applicantsTable.columns.getItem("Applicant").getDataBodyRange().load("values");
            var applicantAgeColumn = applicantsTable.columns.getItem("Age").getDataBodyRange().load("values");
            var applicantGenderColumn = applicantsTable.columns.getItem("Gender").getDataBodyRange().load("values");

            var indexOfSelectedAgent, age, gender;

            // Run all of the above queued-up commands, and return a promise to indicate task completion
            return ctx.sync().then(function () {

                var applicantNameArrays = applicantNameColumn.values;
                var applicantNameColumnValueArray = applicantNameArrays.map(function (item) { return item[0] });

                indexOfSelectedAgent = applicantNameColumnValueArray.indexOf(selectedApplicant);

                var applicantAgeColumnArrays = applicantAgeColumn.values;
                var applicantAgeColumnValueArray = applicantAgeColumnArrays.map(function (item) { return item[0] });

                age = applicantAgeColumnValueArray[indexOfSelectedAgent];

                var applicantGenderColumnArrays = applicantGenderColumn.values;
                var applicantGenderColumnValueArray = applicantGenderColumnArrays.map(function (item) { return item[0] });

                gender = applicantGenderColumnValueArray[indexOfSelectedAgent];

                return ctx.sync();

            })
            .then(function () {
                $('#applicant-age').val(age);
                $('#applicant-gender').val(gender);

            })
        })
            .catch(function (error) {
                handleError(error);
            });
    }

    function fillPolicyRelatedFields(selectedPolicy) {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the transactions table
            var policiesTable = ctx.workbook.tables.getItem("PoliciesTable");
            var policyNameColumn = policiesTable.columns.getItem("Policy").getDataBodyRange().load("values");
            var policyExamColumn = policiesTable.columns.getItem("MedicalExamRequired").getDataBodyRange().load("values");
            var policySampleRateColumn = policiesTable.columns.getItem("SampleRateFor$10000").getDataBodyRange().load("values");

            var indexOfSelectedPolicy, exam, sampleRate;

            // Run all of the above queued-up commands, and return a promise to indicate task completion
            return ctx.sync().then(function () {

                var policyNameArrays = policyNameColumn.values;
                var policyNameColumnValueArray = policyNameArrays.map(function (item) { return item[0] });

                indexOfSelectedPolicy = policyNameColumnValueArray.indexOf(selectedPolicy);

                var policyExamColumnArrays = policyExamColumn.values;
                var policyExamColumnValueArray = policyExamColumnArrays.map(function (item) { return item[0] });

                exam = policyExamColumnValueArray[indexOfSelectedPolicy];

                var policySampleRateColumnArrays = policySampleRateColumn.values;
                var policySampleRateColumnValueArray = policySampleRateColumnArrays.map(function (item) { return item[0] });

                sampleRate = policySampleRateColumnValueArray[indexOfSelectedPolicy];

                return ctx.sync();

            })
            .then(function () {
                $('#exam-required').checked = exam;
                $('#sample-rate').val(sampleRate);
            })
            .then(function () {
                var insuranceAmount = $('#insurance-amount').val();
                var sampleRate = $('#sample-rate').val();
                $('#monthly-payment').text(insuranceAmount * sampleRate / 10000 + '$');
            })
        })

            .catch(function (error) {
                handleError(error);
            });
    }


    function saveProspect() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Create a proxy object for the table rows
            var tableRows = ctx.workbook.tables.getItem('ProspectsTable').rows;
            tableRows.add(null, [[$("#agent-name").val(), $("#applicant-name").val(), $("#applicant-age").val(), $("#applicant-gender").val(), $("#policy-name").val(), "false", $("#sample-rate").val(), $("#insurance-amount").val(), $("#monthly-payment").text()]]);

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
		.catch(function (error) {
		    // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
		    app.showNotification("Error: " + error);
		    console.log("Error: " + error);
		    if (error instanceof OfficeExtension.Error) {
		        console.log("Debug info: " + JSON.stringify(error.debugInfo));
		    }
		});

    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    // Helper function to add and format content in the workbook
    function addContentToWorksheet(sheetObject, rangeAddress, displayText, typeOfText) {

        // Format differently by the type of content
        switch (typeOfText) {
            case "SheetTitle":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 30;
                range.format.font.color = "white";
                range.merge();
                //Fill color in the brand bar
                sheetObject.getRange("A1:M1").format.fill.color = "#41AEBD";
                break;
            case "SheetHeading":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 18;
                range.format.font.color = "#00b3b3";
                range.merge();
                break;
            case "TableHeading":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 12;
                range.format.font.color = "#00b3b3";
                range.merge();
                break;
            case "TableHeaderRow":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                range.format.font.bold = true;
                range.format.font.color = "black";
                break;
            case "TableDataRows":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                sheetObject.getRange(rangeAddress).format.borders.getItem('EdgeBottom').style = 'Continuous';
                sheetObject.getRange(rangeAddress).format.borders.getItem('EdgeTop').style = 'Continuous';
                break;
        }
    }

    // Handle errors
    function handleError(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        app.showNotification("Error: " + error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
})();
/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
