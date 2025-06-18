/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 *  See LICENSE in the source repository root for complete license information.
 */
 
(function () {
    "use strict";
    
    // Define the configurable grading criteria and score values.
    const criteria = ['Content', 'Organization', 'Style', 'Grammar'];
    const score = [0,1,2,3,4,5,6,7,8,9,10];
    const defaultValue = 5;
    
    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            populateScoringDropDowns();
            populatePagePickerDropDown();
            
            // Set up event handlers for the UI.
            $('#getStats').click(getStats);
            $('#addGrade').click(createGrade);
            $('#clear').click(clearGrade);
            $('#openPage').click(openPage);
        });
    };
    
    // Populates the page picker with pages from the current section.
    function populatePagePickerDropDown() {
        OneNote.run(function (context) {
            
            // Get the ID and title of the pages in the current section.
            const pages = context.application.getActiveSection().pages;
            
            // Queue a command to load the id and title for each page.            
            pages.load('id,title');
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    
                    // Get the ID and title of each page, add them as picker options.                  
                    const dropdown = $('#page-picker');
                    $.each(pages.items, function(index, object) {
                        let pageId = object.id;
                        let pageTitle = object.title;
                        
                        if (index === 0) {
                            dropdown.append($('<option selected>').val(pageId).html(pageTitle));
                        } else { 
                            dropdown.append($('<option>').val(pageId).html(pageTitle));
                        }
                    });
                    
                    // Transform the <select> control to an Office UI Fabric dropdown.
                    useFabricDropdown('page-picker-wrapper');
                })
                .catch(function(error) {
                    onError(error);
                }); 
        });
    }
    
    // Get the word and sentence count.
    function getStats() {
        OneNote.run(function (context) {
        
            // Get the collection of pageContent items from the page.
            const pageContents = context.application.getActivePage().contents;
            
            // Queue a command to load the outline property of each pageContent.
            pageContents.load('outline');
            
            // Get the outline on the page.
            // This sample assumes there's only one pageContent on the page with one outline. 
            const pageContent = pageContents.getItem(0);
            
            // Get the paragraphs in the outline.
            const paragraphs = pageContent.outline.paragraphs;
            
            // Queue a command to load the type and richText property of each paragraph.
            paragraphs.load('type,richText');
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    
                    // Get the richText.text from each paragraph.                    
                    let textContent = '';
                    $.each(paragraphs.items, function(index, object) {
                        if (object.type === 'RichText') { 
                            textContent += object.richText.text;
                        }
                    });
                    
                    // Get the word and sentence count and write them to the page.
                    let words = textContent.split(' ');
                    let sentences = textContent.split('. ');                    
                    $('#wordCount').text('Words: ' + words.length);
                    $('#sentenceCount').text('Sentences: ' + sentences.length);
                })                
                .catch(function(error) {
                    onError(error);
                }); 
            });
    }
       
    // Add a table that displays the final grade, individual scores, and comments to the page.
    function addGradeToPage(html) {        
        OneNote.run(function (context) {
            
            // Get the current page.
            const page = context.application.getActivePage();
                       
            // Add an outline with the specified HTML to the page.
            page.addOutline(560, 70, html);
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .catch(function(error) {
                    onError(error);
                }); 
            });
    }
    
    // Open the selected page.
    function openPage() {        
        OneNote.run(function (context) {
            
            // Get the pages in the current section.
            const pages = context.application.getActiveSection().pages;
            
            // Queue a command to load the page collection.            
            pages.load('id');
            
            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    
                    // Get the page with the specified ID from the collection.
                    const selectedPageId = $('#page-picker option:selected').val();
                    let page;
                    $.each(pages.items, function(index, object) {
                        if (object.id === selectedPageId) {
                            page = object;
                            return false;
                        }
                    })
                                        
                    // Open the page in the application.                    
                    context.application.navigateToPage(page);
                    
                    // Run the queued command.
                    return context.sync();
                })
                .catch(function(error) {
                    onError(error);
                }); 
        });
    }
    
    ///* UI only and helpers *///
                  
    // Populates the scoring dropdowns with the score values.
    function populateScoringDropDowns() {
        $.each(criteria, function(index, value) {
            const name = value.toLowerCase();
            const dropdown = $('#' + name);
            $.each(score, function (index) {
                if (index === defaultValue)
                {
                    dropdown.append($('<option selected>').html(index));
                }
                else dropdown.append($('<option>').html(index));
            });
            
            // Transform the <select> control to an Office UI Fabric dropdown.
            useFabricDropdown(name + '-wrapper');
        });
    }
        
    // Get the scores, calculate the grade, and return the results as an HTML table.         
    function createGrade() {        
        let totalScore = 0;
        
        // Create the HTML table that displays the grade. 
        // This string will be passed to the page.addOutline method.
        let table = '<table border=1><tr><td>GRADE</td><td><b>{0}%</b></td></tr>{1}</table>';
        let rows = '';
        
        // Get each score and add it to the running total.
        $.each(criteria, function(index, value) {
            var scoreValue = $('#' + value.toLowerCase()).val();
            var currentScore = parseInt(scoreValue);
            rows += '<tr><td>' + value + '</td><td>' + currentScore + '</td></tr>';
            totalScore = totalScore + currentScore;
        });
        
        // If there's a comment, add it to the table.
        let comments = $('#commentBox').val();
        if (comments) {
            rows += '<tr><td>Comments</td><td><i>' + comments + '</i></td></tr>';
        }
        addGradeToPage(jQuery.validator.format(table, totalScore/criteria.length*10, rows));
    }
    
    // Reset the scoring UI.
    function clearGrade() {
        
        // Reset the dropdowns to the default value.        
        $.each(criteria, function(index, value) {
            $('#' + value.toLowerCase()).val(defaultValue);
        });
        
        // Reset Office UI Fabric dropdowns to the default text.
        $('.ms-Dropdown-title').each(function () {
            if ($(this).parent().attr("id") !== 'page-picker-wrapper') {
                $(this).text(defaultValue);                         
            }
        });
        $('#commentBox').val('');
        $('#wordCount').text('Words:');
        $('#sentenceCount').text('Sentences:');
    }
       
    // Handle errors.
    function onError(error) {
        app.showNotification("Error", "Error: " + error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }   
    
})();
