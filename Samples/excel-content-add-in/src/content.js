/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "./content.css";

(function () {
    "use strict";

    var messageBanner;

    // The onReady function must be run each time a new page is loaded.
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            // Initialize immediately since Office.onReady already waits for DOM.
            var element = document.querySelector('.ms-MessageBanner');
            if (element) {
                try {
                    if (typeof fabric !== 'undefined' && fabric.MessageBanner) {
                        messageBanner = new fabric.MessageBanner(element);
                        messageBanner.hideBanner();
                    }
                } catch (e) {
                    // Handle MessageBanner errors silently.
                }
            }
            
            // Set up event handlers.
            setupEventHandlers();
            
            // Initialize the workbook.
            createProspectsTrackerSheet();
            
            // Import sample data with proper error handling.
            importSampleData().then(() => {
                console.log("Sample data import completed, now calling fillDropDownMenus");
                // Fill dropdown menus after sample data is imported.
                fillDropDownMenus();
            }).catch((error) => {
                console.error("Error importing sample data:", error);
                // Still try to fill dropdown menus even if sample data fails.
                fillDropDownMenus();
            });
        }
    });

    function setupEventHandlers() {
        const applicantNameDropdown = document.getElementById('applicant-name');
        if (applicantNameDropdown) {
            applicantNameDropdown.addEventListener('change', function (e) {
                fillApplicantRelatedFields(this.value);
            });
        }

        const policyNameDropdown = document.getElementById('policy-name');
        if (policyNameDropdown) {
            policyNameDropdown.addEventListener('change', function (e) {
                fillPolicyRelatedFields(this.value);
            });
        }

        const insuranceAmountInput = document.getElementById('insurance-amount');
        if (insuranceAmountInput) {
            insuranceAmountInput.addEventListener('change', function (e) {
                var insuranceAmount = document.getElementById('insurance-amount').value;
                var sampleRate = document.getElementById('sample-rate').value;
                document.getElementById('monthly-payment').textContent = '$' + insuranceAmount * sampleRate / 10000;
            });
        }

        const saveProspectButton = document.getElementById('save-prospect');
        if (saveProspectButton) {
            saveProspectButton.addEventListener('click', saveProspect);
        }

        // Add event handler for notification banner close button.
        const notificationCloseButton = document.querySelector('.ms-MessageBanner-close');
        if (notificationCloseButton) {
            notificationCloseButton.addEventListener('click', function() {
                const notificationBanner = document.getElementById('notificationBanner');
                if (notificationBanner) {
                    notificationBanner.style.display = 'none';
                }
            });
        }
    }

    // Create the sheet to track prospects.
    function createProspectsTrackerSheet() {
        // Run a batch operation against the Excel object model.
        Excel.run(function (ctx) {
            var prospectsSheet = ctx.workbook.worksheets.getActiveWorksheet();
            prospectsSheet.name = "Agent Workspace";

            // Create strings to store all static content to display in the Prospects Tracker sheet.
            var sheetTitle = "Humongous Insurance Agent Center";
            var sheetHeading1 = "In Agent Center, you can easily track and manage prospects.";

            // Add all static content to the Welcome sheet and format the text.
            addContentToWorksheet(prospectsSheet, "B1:X1", sheetTitle, "SheetTitle");
            addContentToWorksheet(prospectsSheet, "B3:K3", sheetHeading1, "SheetHeading");

            // Queue commands to autofit rows and columns in the sheet.
            prospectsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            prospectsSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Run the queued-up commands, and return a promise to indicate task completion.
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Import sample data into tables in the workbook.
    function importSampleData() {
        // Run a batch operation against the Excel object model.
        return Excel.run(function (ctx) {
            // Check if sheets already exist and delete them if they do.
            var worksheets = ctx.workbook.worksheets;
            worksheets.load("items/name");
            
            return ctx.sync().then(() => {
                // Delete existing sheets if they exist.
                const sheetNames = ["Agents", "Applicants", "Policies", "Prospects"];
                sheetNames.forEach(sheetName => {
                    try {
                        const existingSheet = worksheets.items.find(ws => ws.name === sheetName);
                        if (existingSheet) {
                            existingSheet.delete();
                        }
                    } catch (e) {
                        // Sheet doesn't exist, continue.
                    }
                });
                
                return ctx.sync();
            }).then(() => {
                // Queue commands to add a new worksheet to store all the sample data.
                var agentsSheet = ctx.workbook.worksheets.add("Agents");

            // Create strings to store all static content to display in the Agents sheet.
            var sheetTitle = "Humongous Insurance Agent Center";
            var sheetHeading1 = "Agents - Master List";

            // Queue a command to remove gridlines from view.
            agentsSheet.getRange().format.fill.color = "white";

            // Add all static content to the Transactions sheet and format the text.
            addContentToWorksheet(agentsSheet, "B1:K1", sheetTitle, "SheetTitle");
            addContentToWorksheet(agentsSheet, "B3:K3", sheetHeading1, "SheetHeading");

            // Queue a command to add a new table.
            var agentsTable = ctx.workbook.tables.add('Agents!B6:B6', true);
            agentsTable.name = "AgentsTable";

            // Queue a command to set the header row.
            agentsTable.getHeaderRowRange().values = [["AgentName"]];
            var tableRows = agentsTable.rows;

            tableRows.add(null, [["Aanandini Kidambi"]]);
            tableRows.add(null, [["Jordan Hopkins"]]);
            tableRows.add(null, [["Amelie Laffer"]]);
            tableRows.add(null, [["Ya-ting Lo"]]);
            tableRows.add(null, [["Chelsea Leigh"]]);
            tableRows.add(null, [["Badanika Atluri"]]);

            // Queue commands to format the table.
            addContentToWorksheet(agentsSheet, "B6:B6", "", "TableHeaderRow");
            addContentToWorksheet(agentsSheet, "B7:B12", "", "TableDataRows");

            // Queue commands to auto-fit columns and rows.
            agentsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            agentsSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Applicants sheet.
            var applicantsSheet = ctx.workbook.worksheets.add("Applicants");

            // Create strings to store all static content to display in the Applicants sheet.
            var sheetTitle = "Humongous Insurance Agent Center";
            var sheetHeading1 = "Applicants - Master List";

            // Queue a command to remove gridlines from view.
            applicantsSheet.getRange().format.fill.color = "white";

            // Add all static content to the Applicants sheet and format the text.
            addContentToWorksheet(applicantsSheet, "B1:K1", sheetTitle, "SheetTitle");
            addContentToWorksheet(applicantsSheet, "B3:K3", sheetHeading1, "SheetHeading");

            // Queue a command to add a new table.
            var applicantsTable = ctx.workbook.tables.add('Applicants!B6:D6', true);
            applicantsTable.name = "ApplicantsTable";

            // Queue a command to set the header row.
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

            // Format the table header and data rows.
            addContentToWorksheet(applicantsSheet, "B6:D6", "", "TableHeaderRow");
            addContentToWorksheet(applicantsSheet, "B7:D18", "", "TableDataRows");

            // Queue commands to auto-fit columns and rows.
            applicantsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            applicantsSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Policies sheet.
            var policiesSheet = ctx.workbook.worksheets.add("Policies");

            // Create strings to store all static content to display in the Policies sheet.
            var sheetTitle = "Humongous Insurance Agent Center";
            var sheetHeading1 = "Policies - Master List";

            // Queue a command to remove gridlines from view.
            policiesSheet.getRange().format.fill.color = "white";

            // Add all static content to the Policies sheet and format the text.
            addContentToWorksheet(policiesSheet, "B1:K1", sheetTitle, "SheetTitle");
            addContentToWorksheet(policiesSheet, "B3:K3", sheetHeading1, "SheetHeading");

            // Queue a command to add a new table.
            var policiesTable = ctx.workbook.tables.add('Policies!B6:D6', true);
            policiesTable.name = "PoliciesTable";

            // Queue a command to set the header row.
            policiesTable.getHeaderRowRange().values = [["PolicyName", "Medical Exam Required", "Sample Rate For $10000"]];
            var tableRows = policiesTable.rows;

            tableRows.add(null, [["Whole Life", "No", "$42.5"]]);
            tableRows.add(null, [["Universal Life", "No", "$35.5"]]);
            tableRows.add(null, [["Variable Life", "Yes", "$28.25"]]);
            tableRows.add(null, [["Term Life", "Yes", "$4.75"]]);

            // Format the table header and data rows.
            addContentToWorksheet(policiesSheet, "B6:D6", "", "TableHeaderRow");
            addContentToWorksheet(policiesSheet, "B7:D10", "", "TableDataRows");

            // Queue commands to auto-fit columns and rows.
            policiesSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            policiesSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Add a new sheet to save prospects.
            var savedProspectsSheet = ctx.workbook.worksheets.add("Prospects");

            // Add the prospects table.
            var prospectsTable = ctx.workbook.tables.add('Prospects!A1:I1', true);
            prospectsTable.name = "ProspectsTable";
            prospectsTable.getHeaderRowRange().values = [["Agent", "Applicant", "Age", "Gender", "Policy", "Exam Required", "Sample Rate", "Insurance Amount", "Monthly Payment"]];

            // Run the queued-up commands, and return a promise to indicate task completion.
            return ctx.sync().then(() => {
                console.log("Sample data import completed successfully");
            });
            }); // Close the .then() chain.
        })
        .catch(function (error) {
            console.error("Error in importSampleData:", error);
            handleError(error);
            throw error; // Re-throw to allow calling function to handle it.
        });
    }

    function fillDropDownMenus() {
        console.log("fillDropDownMenus called");
        // Run a batch operation against the Excel object model.
        Excel.run(function (ctx) {
            console.log("Excel.run started in fillDropDownMenus");
            // Queue a command to get the agents table.
            var agentsTable = ctx.workbook.tables.getItem("AgentsTable");
            var agentNameColumn = agentsTable.columns.getItem("AgentName").getDataBodyRange().load("values");

            // Queue a command to get the applicants table.
            var applicantsTable = ctx.workbook.tables.getItem("ApplicantsTable");
            var applicantNameColumn = applicantsTable.columns.getItem("Applicant").getDataBodyRange().load("values");

            // Queue a command to get the policies table.
            var policiesTable = ctx.workbook.tables.getItem("PoliciesTable");
            var policyNameColumn = policiesTable.columns.getItem("PolicyName").getDataBodyRange().load("values");

            console.log("About to sync and load table data");
            // Run all of the above queued-up commands, and return a promise to indicate task completion.
            return ctx.sync().then(function () {
                console.log("Tables loaded successfully");
                var agentNames = agentNameColumn.values;
                var agentDropdown = document.getElementById("agent-name");
                console.log("Agent names:", agentNames);
                console.log("Agent dropdown element:", agentDropdown);

                // Clear existing options except the first one.
                while (agentDropdown.options.length > 1) {
                    agentDropdown.removeChild(agentDropdown.lastChild);
                }

                // For each agent name, add a dropdown item in the UI.
                for (var i = 0; i < agentNames.length; i++) {
                    var name = agentNames[i][0];
                    var newOption = document.createElement('option');
                    newOption.value = name;
                    newOption.textContent = name;
                    agentDropdown.appendChild(newOption);
                    console.log("Added agent option:", name);
                }

                var applicantNames = applicantNameColumn.values;
                var applicantDropdown = document.getElementById("applicant-name");
                console.log("Applicant names:", applicantNames);

                // Clear existing options except the first one.
                while (applicantDropdown.options.length > 1) {
                    applicantDropdown.removeChild(applicantDropdown.lastChild);
                }

                // For each applicant name, add a dropdown item in the UI.
                for (var i = 0; i < applicantNames.length; i++) {
                    var name = applicantNames[i][0];
                    var newOption = document.createElement('option');
                    newOption.value = name;
                    newOption.textContent = name;
                    applicantDropdown.appendChild(newOption);
                    console.log("Added applicant option:", name);
                }

                var policyNames = policyNameColumn.values;
                var policyDropdown = document.getElementById("policy-name");
                console.log("Policy names:", policyNames);

                // Clear existing options except the first one.
                while (policyDropdown.options.length > 1) {
                    policyDropdown.removeChild(policyDropdown.lastChild);
                }

                // For each policy name, add a dropdown item in the UI.
                for (var i = 0; i < policyNames.length; i++) {
                    var name = policyNames[i][0];
                    var newOption = document.createElement('option');
                    newOption.value = name;
                    newOption.textContent = name;
                    policyDropdown.appendChild(newOption);
                    console.log("Added policy option:", name);
                }
                
                console.log("Finished populating all dropdowns");
            });
        })
        .catch(function (error) {
            console.error("Error in fillDropDownMenus:", error);
            handleError(error);
        });
    }

    function fillApplicantRelatedFields(selectedApplicant) {
        // Run a batch operation against the Excel object model.
        Excel.run(function (ctx) {
            // Queue a command to get the applicants table.
            var applicantsTable = ctx.workbook.tables.getItem("ApplicantsTable");
            var applicantNameColumn = applicantsTable.columns.getItem("Applicant").getDataBodyRange().load("values");
            var applicantAgeColumn = applicantsTable.columns.getItem("Age").getDataBodyRange().load("values");
            var applicantGenderColumn = applicantsTable.columns.getItem("Gender").getDataBodyRange().load("values");

            var indexOfSelectedAgent, age, gender;

            // Run all of the above queued-up commands, and return a promise to indicate task completion.
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
                document.getElementById('applicant-age').value = age;
                document.getElementById('applicant-gender').value = gender;
            })
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    function fillPolicyRelatedFields(selectedPolicy) {
        // Run a batch operation against the Excel object model.
        Excel.run(function (ctx) {
            // Queue a command to get the policies table.
            var policiesTable = ctx.workbook.tables.getItem("PoliciesTable");
            var policyNameColumn = policiesTable.columns.getItem("PolicyName").getDataBodyRange().load("values");
            var policyExamColumn = policiesTable.columns.getItem("Medical Exam Required").getDataBodyRange().load("values");
            var policySampleRateColumn = policiesTable.columns.getItem("Sample Rate For $10000").getDataBodyRange().load("values");

            var indexOfSelectedPolicy, exam, sampleRate;

            // Run all of the above queued-up commands, and return a promise to indicate task completion.
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
                document.getElementById('exam-required').checked = exam === "Yes";
                document.getElementById('sample-rate').value = sampleRate;
            })
            .then(function () {
                var insuranceAmount = document.getElementById('insurance-amount').value;
                var sampleRate = document.getElementById('sample-rate').value;
                document.getElementById('monthly-payment').textContent = '$' + insuranceAmount * sampleRate / 10000;
            })
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    function saveProspect() {
        // Run a batch operation against the Excel object model.
        Excel.run(function (ctx) {
            // Create a proxy object for the table rows.
            var tableRows = ctx.workbook.tables.getItem('ProspectsTable').rows;
            tableRows.add(null, [[
                document.getElementById("agent-name").value, 
                document.getElementById("applicant-name").value, 
                document.getElementById("applicant-age").value, 
                document.getElementById("applicant-gender").value, 
                document.getElementById("policy-name").value, 
                document.getElementById('exam-required').checked, 
                document.getElementById("sample-rate").value, 
                document.getElementById("insurance-amount").value, 
                document.getElementById("monthly-payment").textContent
            ]]);

            // Run the queued-up commands, and return a promise to indicate task completion.
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Helper function for displaying notifications.
    function showNotification(header, content) {
        document.getElementById("notificationHeader").textContent = header;
        document.getElementById("notificationBody").textContent = content;
        
        // Show the notification banner.
        const notificationBanner = document.getElementById("notificationBanner");
        if (notificationBanner) {
            notificationBanner.style.display = 'block';
        }
        
        if (messageBanner) {
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    }

    // Helper function to add and format content in the workbook.
    function addContentToWorksheet(sheetObject, rangeAddress, displayText, typeOfText) {
        // Format differently by the type of content.
        switch (typeOfText) {
            case "SheetTitle":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 30;
                range.format.font.color = "white";
                range.merge();
                // Fill color in the brand bar.
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

    // Handle errors.
    function handleError(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution.
        showNotification("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            // Silent debug info handling.
        }
    }

    // Make functions available globally if needed.
    window.showNotification = showNotification;
})();
