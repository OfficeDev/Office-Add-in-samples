/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

/// <reference types="office-js-preview" />

/**
 * Basic function to show how to insert a table into a Word document.
 */
console.log("Loading Weather.razor.ts");

export async function createWeatherTable(forecasts: any[]) {

    console.log("We are now entering function: createWeatherTable");
    console.log("Received forecasts:", forecasts);

    try {
        await Word.run(async function (context) {
            // Get the current document body
            const body = context.document.body;

            // Insert a title
            const title = body.insertParagraph("Weather Forecast", Word.InsertLocation.end);
            title.styleBuiltIn = Word.BuiltInStyleName.title;

            // Create table data: header row + data rows
            const table = body.insertTable(
                forecasts.length + 1, 
                4, 
                Word.InsertLocation.end, 
                [
                    ["Date", "Temp (°C)", "Temp (°F)", "Summary"],
                    ...forecasts.map(f => [
                        f.date,
                        f.temperatureC.toString(),
                        f.temperatureF.toString(),
                        f.summary || ""
                    ])
                ]
            );
            table.styleBuiltIn = Word.BuiltInStyleName.gridTable5Dark_Accent1;

            // Add some spacing after the table
            body.insertParagraph("", Word.InsertLocation.end);

            // Single sync to commit all changes
            await context.sync();
            console.log("Weather Forecast table created successfully.");
        });
    } catch (error) {
        console.error("Error creating Weather Forecast table: ", error);
        throw error;
    }
}