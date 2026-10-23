let visualization = (function () {
    'use strict';

    let visualization = {};

    // Generates and returns an Office.TableData object with sample data.
    visualization.generateSampleData = function () {
        const sampleHeaders = [['Name', 'Grade']];
        const sampleRows = [
            ['Ben', 79],
            ['Amy', 95],
            ['Jacob', 86],
            ['Ernie', 93]];
        return new Office.TableData(sampleRows, sampleHeaders);
    };

    // Displays a visualization based on the following parameters:
    //        element:  An element where the visualization will be displayed.
    //        data:  An Office.TableData object that contains the data.
    //        errorHandler:  An error callback that accepts a string description.
    visualization.display = function (element, data, errorHandler) {
        if (data.rows.length < 1 || data.rows[0].length < 2) {
            errorHandler('The data range must contain at least 1 row and at least 2 columns.');
            return;
        }

        const maxBarWidthInPixels = 200;
        let table = document.createElement('table');
        table.className = "visualization";

        if (data.headers !== null && data.headers.length > 0) {
            let headerRow = document.createElement('tr');
            table.appendChild(headerRow);
            let th1 = document.createElement('th');
            th1.textContent = data.headers[0][0];
            headerRow.appendChild(th1);
            let th2 = document.createElement('th');
            th2.textContent = data.headers[0][1];
            headerRow.appendChild(th2);
        }

        for (let i = 0; i < data.rows.length; i++) {
            let row = document.createElement('tr');
            table.appendChild(row);
            let column1 = document.createElement('td');
            row.appendChild(column1);
            let column2 = document.createElement('td');
            row.appendChild(column2);

            column1.textContent = data.rows[i][0];
            let value = data.rows[i][1];
            let width = maxBarWidthInPixels * value / 100.0;
            let visualizationBar = document.createElement('div');
            column2.appendChild(visualizationBar);
            visualizationBar.className = 'bar';
            visualizationBar.style.width = width + 'px';
            visualizationBar.textContent = value;
        }

        element.innerHTML = table.outerHTML;
    };

    return visualization;
})();
