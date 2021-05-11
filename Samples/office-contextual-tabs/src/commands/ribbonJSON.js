/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//The following function defines and returns the JSON object that describes the contextual tabs.
//Make changes to the JSON text if you want to modify the contextual tabs.
function getContextualRibbonJSON() {
    const sourceUrl = "https://officedev.github.io/PnP-OfficeAddins/Samples/office-contextual-tabs/";
  return {
    "actions": [
        {
            "id": "idRibbonAction",
            "type": "ExecuteFunction",
            "functionName": "runRibbonAction"
        },
        {
            "id": "showTaskpaneContoso",
            "type": "ShowTaskpane",
            "title": "Contoso task pane",
            "supportPinning": false
          }
    ],
    "tabs": [
        {
            "id": "tabTableData",
            "label": "Table Data",
            "visible": false,
            "groups": [
                {
                    "id": "grpSources",
                    "label": "Data sources",
                    "icon": [
                        {
                            "size": 16,
                            "sourceLocation": sourceUrl + "assets/icon-16.png"
                        },
                        {
                            "size": 32,
                            "sourceLocation": sourceUrl + "assets/icon-32.png"
                        },
                        {
                            "size": 80,
                            "sourceLocation": sourceUrl + "assets/icon-80.png"
                        }
                    ],
                    "controls": [
                        {
                            "type": "Menu",
                            "id": "mnuDataSources",
                            "label": "Import data",
                            "toolTip": "Select data source to import from",
                            "superTip": {
                                "title": "Data Sources",
                                "description": "Select data source to import from."
                            },
                            "icon": [
                                {
                                    "size": 16,
                                    "sourceLocation": sourceUrl + "assets/icon-16.png"
                                },
                                {
                                    "size": 32,
                                    "sourceLocation": sourceUrl + "assets/icon-32.png"
                                },
                                {
                                    "size": 80,
                                    "sourceLocation": sourceUrl + "assets/icon-80.png"
                                }
                            ],
                            "items": [
                                {
                                    "type": "MenuItem",
                                    "id": "itmExternalExcel",
                                    "enabled": true,
                                    "icon": [
                                        {
                                            "size": 16,
                                            "sourceLocation": sourceUrl + "assets/icon-16.png"
                                        },
                                        {
                                            "size": 32,
                                            "sourceLocation": sourceUrl + "assets/icon-32.png"
                                        },
                                        {
                                            "size": 80,
                                            "sourceLocation": sourceUrl + "assets/icon-80.png"
                                        }
                                    ],
                                    "label": "External Excel file",
                                    "toolTip": "Sync with external Excel file",
                                    "superTip": {
                                        "title": "External Excel file",
                                        "description": "Sync with external Excel file"
                                    },
                                    "actionId": "idRibbonAction"
                                },
                                {
                                    "type": "MenuItem",
                                    "id": "itmSQLSource",
                                    "enabled": true,
                                    "icon": [
                                        {
                                            "size": 16,
                                            "sourceLocation": sourceUrl + "assets/icon-16.png"
                                        },
                                        {
                                            "size": 32,
                                            "sourceLocation": sourceUrl + "assets/icon-32.png"
                                        },
                                        {
                                            "size": 80,
                                            "sourceLocation": sourceUrl + "assets/icon-80.png"
                                        }
                                    ],
                                    "label": "SQL Database",
                                    "toolTip": "Sync with SQL Database",
                                    "superTip": {
                                        "title": "SQL Database",
                                        "description": "Sync with SQL Database."
                                    },
                                    "actionId": "idRibbonAction"
                                }
                            ]
                        }
                    ]
                },
                {
                    "id": "grpData",
                    "label": "Sync data",
                    "icon": [
                        {
                            "size": 16,
                            "sourceLocation": sourceUrl + "assets/icon-16.png"
                        },
                        {
                            "size": 32,
                            "sourceLocation": sourceUrl + "assets/icon-32.png"
                        },
                        {
                            "size": 80,
                            "sourceLocation": sourceUrl + "assets/icon-80.png"
                        }
                    ],
                    "controls": [
                        {
                            "type": "Button",
                            "id": "btnRefresh",
                            "enabled": false,
                            "icon": [
                                {
                                    "size": 16,
                                    "sourceLocation": sourceUrl + "assets/icon-16.png"
                                },
                                {
                                    "size": 32,
                                    "sourceLocation": sourceUrl + "assets/icon-32.png"
                                },
                                {
                                    "size": 80,
                                    "sourceLocation": sourceUrl + "assets/icon-80.png"
                                }
                            ],
                            "label": "Refresh",
                            "toolTip": "Refresh table with latest data from data source",
                            "superTip": {
                                "title": "Refresh table data",
                                "description": "Refresh table with latest data from data source."
                            },
                            "actionId": "idRibbonAction"
                        },
                        {
                            "type": "Button",
                            "id": "btnSubmit",
                            "enabled": false,
                            "icon": [
                                {
                                    "size": 16,
                                    "sourceLocation": sourceUrl + "assets/icon-16.png"
                                },
                                {
                                    "size": 32,
                                    "sourceLocation": sourceUrl + "assets/icon-32.png"
                                },
                                {
                                    "size": 80,
                                    "sourceLocation": sourceUrl + "assets/icon-80.png"
                                }
                            ],
                            "label": "Submit",
                            "toolTip": "Submit data changes from table to data source",
                            "superTip": {
                                "title": "Submit",
                                "description": "Submit data changes from table to data source."
                            },
                            "actionId": "idRibbonAction"
                        }
                    ]
                },
                {
                    "id": "grpTaskpane",
                    "label": "Task pane",
                    "icon": [
                        {
                            "size": 16,
                            "sourceLocation": sourceUrl + "assets/icon-16.png"
                        },
                        {
                            "size": 32,
                            "sourceLocation": sourceUrl + "assets/icon-32.png"
                        },
                        {
                            "size": 80,
                            "sourceLocation": sourceUrl + "assets/icon-80.png"
                        }
                    ],
                    "controls": [
                        {
                            "type": "Button",
                            "id": "btnShowTaskPane",
                            "enabled": true,
                            "icon": [
                                {
                                    "size": 16,
                                    "sourceLocation": sourceUrl + "assets/icon-16.png"
                                },
                                {
                                    "size": 32,
                                    "sourceLocation": sourceUrl + "assets/icon-32.png"
                                },
                                {
                                    "size": 80,
                                    "sourceLocation": sourceUrl + "assets/icon-80.png"
                                }
                            ],
                            "label": "Show task pane",
                            "toolTip": "Show Contoso task pane",
                            "superTip": {
                                "title": "Show task pane",
                                "description": "Show Contoso task pane."
                            },
                            "actionId": "showTaskpaneContoso"
                        }
                    ]
                }
            ]
        }
    ]
};
}
