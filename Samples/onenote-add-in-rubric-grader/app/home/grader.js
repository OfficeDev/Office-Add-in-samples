// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
 
(function () {
    "use strict";
    
    // Define the configurable grading criteria and score values.
    const criteria = ['Content', 'Organization', 'Style', 'Grammar'];
    const score = [0,1,2,3,4,5,6,7,8,9,10];
    const defaultValue = 5;
    
    // Initialize the add-in when the page loads.
    Office.onReady(() => {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initializeApp);
        } else {
            initializeApp();
        }
    });
    
    function initializeApp() {
        console.log('Initializing app...');
        app.initialize();

        console.log('Populating dropdowns...');
        populateScoringDropDowns();
        populatePagePickerDropDown();
        
        // Set up event handlers for the UI controls.
        console.log('Setting up event handlers...');
        document.getElementById('getStats').addEventListener('click', getStats);
        document.getElementById('addGrade').addEventListener('click', createGrade);
        document.getElementById('clear').addEventListener('click', clearGrade);
        document.getElementById('openPage').addEventListener('click', openPage);
        console.log('Initialization complete.');
    }
    
    // Populate the page picker with pages from the current section.
    async function populatePagePickerDropDown() {
        try {
            await OneNote.run(async (context) => {
                
                // Get the ID and title of the pages in the current section.
                const pages = context.application.getActiveSection().pages;
                
                // Load the id and title for each page.
                pages.load('id,title');
                await context.sync();
                
                // Add each page as an option in the dropdown.
                const dropdown = document.getElementById('page-picker');
                pages.items.forEach((object, index) => {
                    const pageId = object.id;
                    const pageTitle = object.title;
                    
                    const option = document.createElement('option');
                    option.value = pageId;
                    option.textContent = pageTitle;
                    
                    if (index === 0) {
                        option.selected = true;
                    }
                    
                    dropdown.appendChild(option);
                });
            });
        } catch (error) {
            onError(error);
        }
    }
    
    // Get the word and sentence count from the current page.
    async function getStats() {
        try {
            await OneNote.run(async (context) => {
            
                // Get the collection of pageContent items from the page.
                const pageContents = context.application.getActivePage().contents;
                
                // Load the outline property of each pageContent.
                pageContents.load('outline');
                
                // Get the outline on the page.
                // This sample assumes there's only one pageContent on the page with one outline. 
                const pageContent = pageContents.getItem(0);
                
                // Get the paragraphs in the outline.
                const paragraphs = pageContent.outline.paragraphs;
                
                // Load the type and richText property of each paragraph.
                paragraphs.load('type,richText');
                await context.sync();
                
                // Get the text content from each rich text paragraph.
                let textContent = '';
                paragraphs.items.forEach((object) => {
                    if (object.type === 'RichText') { 
                        textContent += object.richText.text;
                    }
                });
                
                // Calculate word and sentence counts and display them.
                const words = textContent.split(' ');
                const sentences = textContent.split('. ');                    
                document.getElementById('wordCount').textContent = 'Words: ' + words.length;
                document.getElementById('sentenceCount').textContent = 'Sentences: ' + sentences.length;
            });
        } catch (error) {
            onError(error);
        }
    }
       
    // Add a grading table to the current page.
    async function addGradeToPage(html) {        
        try {
            await OneNote.run(async (context) => {
                
                // Get the current page.
                const page = context.application.getActivePage();
                           
                // Add an outline with the specified HTML to the page.
                page.addOutline(560, 70, html);
                await context.sync();
            });
        } catch (error) {
            onError(error);
        }
    }
    
    // Open the selected page in OneNote.
    async function openPage() {        
        try {
            await OneNote.run(async (context) => {
                
                // Get the pages in the current section.
                const pages = context.application.getActiveSection().pages;
                
                // Load the page collection.
                pages.load('id');
                await context.sync();
                
                // Find the page with the specified ID.
                const selectedPageElement = document.querySelector('#page-picker option:checked');
                const selectedPageId = selectedPageElement ? selectedPageElement.value : null;
                let page;
                pages.items.forEach((object) => {
                    if (object.id === selectedPageId) {
                        page = object;
                    }
                });
                                    
                // Navigate to the selected page.
                context.application.navigateToPage(page);
                await context.sync();
            });
        } catch (error) {
            onError(error);
        }
    }
    
    ///* UI helpers *///
                  
    // Populate the scoring dropdowns with score values.
    function populateScoringDropDowns() {
        console.log('Populating scoring dropdowns...');
        criteria.forEach((value) => {
            const name = value.toLowerCase();
            const dropdown = document.getElementById(name);
            
            if (!dropdown) {
                console.error(`Dropdown element not found: ${name}`);
                return;
            }
            
            console.log(`Populating dropdown: ${name}`);
            
            // Clear existing options first
            dropdown.innerHTML = '';
            
            score.forEach((scoreValue, index) => {
                const option = document.createElement('option');
                option.value = index;
                option.textContent = index;
                
                if (index === defaultValue) {
                    option.selected = true;
                }
                
                dropdown.appendChild(option);
            });
            
            console.log(`Dropdown ${name} now has ${dropdown.options.length} options`);
        });
    }
        
    // Calculate the grade and create an HTML table with the results.
    function createGrade() {        
        let totalScore = 0;
        
        // Create the HTML table that displays the grade. 
        // This string will be passed to the addGradeToPage method.
        const table = '<table border=1><tr><td>GRADE</td><td><b>{0}%</b></td></tr>{1}</table>';
        let rows = '';
        
        // Get each score and add it to the running total.
        criteria.forEach((value) => {
            const scoreElement = document.getElementById(value.toLowerCase());
            const scoreValue = scoreElement.value;
            const currentScore = parseInt(scoreValue);
            rows += '<tr><td>' + value + '</td><td>' + currentScore + '</td></tr>';
            totalScore = totalScore + currentScore;
        });
        
        // Add the comment to the table if one was provided.
        const comments = document.getElementById('commentBox').value;
        if (comments) {
            rows += '<tr><td>Comments</td><td><i>' + comments + '</i></td></tr>';
        }
        
        // Use string replacement to format the final table.
        const finalTable = table.replace('{0}', (totalScore/criteria.length*10)).replace('{1}', rows);
        addGradeToPage(finalTable);
    }
    
    // Reset the scoring UI to default values.
    function clearGrade() {
        
        // Reset the dropdowns to the default value.
        criteria.forEach((value) => {
            document.getElementById(value.toLowerCase()).value = defaultValue;
        });
        
        document.getElementById('commentBox').value = '';
        document.getElementById('wordCount').textContent = 'Words:';
        document.getElementById('sentenceCount').textContent = 'Sentences:';
    }
       
    // Handle errors and display them to the user.
    function onError(error) {
        app.showNotification("Error", "Error: " + error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }   
    
})();
