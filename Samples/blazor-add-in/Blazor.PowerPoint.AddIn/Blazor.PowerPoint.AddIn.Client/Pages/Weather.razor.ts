/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

/**
 * Function to create a new slide with a weather data table in a PowerPoint Presentation.
 * Uses the PowerPoint JavaScript Table API (PowerPointApi 1.8+).
 */
console.log("Loading Weather.razor.ts");

interface WeatherRow {
    date: string;
    tempC: string;
    tempF: string;
    summary: string;
}

export async function createWeatherSlide(tableData: WeatherRow[]) {

    console.log("We are now entering function: createWeatherSlide");

    try {
        await PowerPoint.run(async function (context) {

            // Create a new slide
            context.presentation.slides.add();
            await context.sync();

            // Navigate to the new slide
            await goToLastSlide();
            console.log("Slide created successfully.");

            var slides = context.presentation.slides;
            slides.load("items");

            // Load slide dimensions for centering
            const pageSetup = context.presentation.pageSetup;
            pageSetup.load("slideWidth, slideHeight");
            await context.sync();

            console.log(slides.items.length);

            const newSlide = slides.getItemAt(slides.items.length - 1);
            const shapes: PowerPoint.ShapeCollection = newSlide.shapes;

            // Remove default placeholder shapes to get a blank slide
            await removeSlidePlaceholders(shapes);

            // Add a weather-themed background image (added first so it's behind other shapes)
            const bgBase64 = generateWeatherBackground(
                Math.round(pageSetup.slideWidth * 2),
                Math.round(pageSetup.slideHeight * 2)
            );
            shapes.addPicture(bgBase64, {
                left: 0,
                top: 0,
                width: pageSetup.slideWidth,
                height: pageSetup.slideHeight
            });
            await context.sync();

            // Add a centered title textbox
            const titleWidth = 600;
            const titleHeight = 40;
            const titleTop = 20;
            var textbox: PowerPoint.Shape = shapes.addTextBox("Weather Forecast",
                {
                    left: (pageSetup.slideWidth - titleWidth) / 2,
                    top: titleTop,
                    height: titleHeight,
                    width: titleWidth
                });
            textbox.name = "WeatherForecastTitle";
            textbox.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;

            // Build the 2D values array: header row + data rows
            const headerRow = ["Date", "Temp (C)", "Temp (F)", "Summary"];
            const values: string[][] = [headerRow];
            for (const row of tableData) {
                values.push([row.date, row.tempC, row.tempF, row.summary]);
            }

            const rowCount = values.length;
            const columnCount = headerRow.length;

            // Create specificCellProperties 2D array for header row styling
            const specificCellProperties: PowerPoint.TableCellProperties[][] =
                Array(rowCount).fill("").map(_ =>
                    Array(columnCount).fill("").map(_ => ({} as PowerPoint.TableCellProperties))
                );

            // Style each header cell with bold white text on dark background
            const headerRowProps = specificCellProperties[0];
            if (headerRowProps) {
                for (let col = 0; col < columnCount; col++) {
                    headerRowProps[col] = {
                        fill: { color: "#0F6CBD" },
                        font: { bold: true, color: "white", size: 12 },
                        horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment.center
                    };
                }
            }

            // Style data rows with center alignment for numeric columns
            for (let row = 1; row < rowCount; row++) {
                const dataRow = specificCellProperties[row];
                if (dataRow) {
                    dataRow[1] = {
                        horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment.center
                    };
                    dataRow[2] = {
                        horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment.center
                    };
                }
            }

            // Add the table to the slide using the PowerPoint Table API
            const tableOptions: PowerPoint.TableAddOptions = {
                values: values,
                uniformCellProperties: {
                    font: { size: 11 },
                    borders: {
                        bottom: { color: "#E0E0E0", weight: 1 },
                        left: { color: "#E0E0E0", weight: 1 },
                        right: { color: "#E0E0E0", weight: 1 },
                        top: { color: "#E0E0E0", weight: 1 }
                    }
                },
                specificCellProperties: specificCellProperties,
                columns: [
                    { columnWidth: 120 },
                    { columnWidth: 100 },
                    { columnWidth: 100 },
                    { columnWidth: 160 }
                ]
            };

            const tableShape = shapes.addTable(rowCount, columnCount, tableOptions);

            // Load the table shape dimensions to calculate centered position
            tableShape.load("width, height");
            await context.sync();

            // Center the table horizontally and vertically on the slide
            // Account for the title area (titleTop + titleHeight + margin)
            const titleAreaBottom = titleTop + titleHeight + 10;
            const availableHeight = pageSetup.slideHeight - titleAreaBottom;
            tableShape.left = (pageSetup.slideWidth - tableShape.width) / 2;
            tableShape.top = titleAreaBottom + (availableHeight - tableShape.height) / 2;
            await context.sync();

            console.log("Weather Forecast table slide created successfully.");
        });
    } catch (error) {
        console.error("Error creating Weather Forecast slide: ", error);
    }
}

/**
 * Generates a weather-themed background image using the Canvas API.
 * Returns the raw base64 PNG data (without the data URI prefix).
 */
function generateWeatherBackground(widthPx: number, heightPx: number): string {
    const canvas = document.createElement("canvas");
    canvas.width = widthPx;
    canvas.height = heightPx;
    const ctx = canvas.getContext("2d")!;

    // Sky gradient
    const skyGradient = ctx.createLinearGradient(0, 0, 0, heightPx);
    skyGradient.addColorStop(0, "#4A90D9");
    skyGradient.addColorStop(0.5, "#87CEEB");
    skyGradient.addColorStop(0.85, "#B0E0E6");
    skyGradient.addColorStop(1, "#E8F4FD");
    ctx.fillStyle = skyGradient;
    ctx.fillRect(0, 0, widthPx, heightPx);

    // Sun with glow
    const sunX = widthPx * 0.82;
    const sunY = heightPx * 0.18;
    const sunRadius = widthPx * 0.04;

    const sunGlow = ctx.createRadialGradient(sunX, sunY, sunRadius * 0.5, sunX, sunY, sunRadius * 3);
    sunGlow.addColorStop(0, "rgba(255, 236, 140, 0.5)");
    sunGlow.addColorStop(0.5, "rgba(255, 223, 100, 0.15)");
    sunGlow.addColorStop(1, "rgba(255, 223, 100, 0)");
    ctx.fillStyle = sunGlow;
    ctx.fillRect(0, 0, widthPx, heightPx);

    ctx.beginPath();
    ctx.arc(sunX, sunY, sunRadius, 0, Math.PI * 2);
    ctx.fillStyle = "rgba(255, 223, 100, 0.7)";
    ctx.fill();

    // Clouds
    drawCloud(ctx, widthPx * 0.1, heightPx * 0.12, widthPx * 0.05, "rgba(255, 255, 255, 0.5)");
    drawCloud(ctx, widthPx * 0.35, heightPx * 0.08, widthPx * 0.06, "rgba(255, 255, 255, 0.45)");
    drawCloud(ctx, widthPx * 0.6, heightPx * 0.22, widthPx * 0.04, "rgba(255, 255, 255, 0.35)");
    drawCloud(ctx, widthPx * 0.2, heightPx * 0.3, widthPx * 0.035, "rgba(255, 255, 255, 0.25)");
    drawCloud(ctx, widthPx * 0.85, heightPx * 0.38, widthPx * 0.03, "rgba(255, 255, 255, 0.2)");

    // Extract raw base64 data (strip the "data:image/png;base64," prefix)
    const dataUrl = canvas.toDataURL("image/png");
    return dataUrl.substring(dataUrl.indexOf(",") + 1);
}

/**
 * Draws a soft cloud shape at the given position and scale.
 */
function drawCloud(ctx: CanvasRenderingContext2D, x: number, y: number, scale: number, color: string) {
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.arc(x, y, scale, 0, Math.PI * 2);
    ctx.arc(x + scale * 0.8, y - scale * 0.4, scale * 0.85, 0, Math.PI * 2);
    ctx.arc(x + scale * 1.7, y, scale, 0, Math.PI * 2);
    ctx.arc(x + scale * 0.85, y + scale * 0.2, scale * 0.7, 0, Math.PI * 2);
    ctx.fill();
}