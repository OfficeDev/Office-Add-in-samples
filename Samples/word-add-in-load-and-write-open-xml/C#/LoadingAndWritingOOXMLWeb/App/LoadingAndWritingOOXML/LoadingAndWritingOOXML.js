/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {
        // If setSelectedDataAsync method is supported by the host application
        // the UI buttons are hooked up to call the method else the buttons are removed

        if (Office.context.document.setSelectedDataAsync) {

            clickHandler();

        }

        else {
            $('#setFText').remove();
            $('#setSText').remove();
            $('#setImage').remove();
            $('#setBox').remove();
            $('#setShape').remove();
            $('#setControl').remove();
            $('#setFTable').remove();
            $('#setSTable').remove();
            $('#setSmartArt').remove();
            $('#setChart').remove();
        }
    });
};


function writeContent(fileName) {

    const myOOXMLRequest = new XMLHttpRequest();
    let myXML;
    myOOXMLRequest.open('GET', fileName, false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });

}

function writeMarkup(fileName) {

    const myOOXMLRequest = new XMLHttpRequest();
    let myXML;

    //Set the format for the markup
    myOOXMLRequest.open('GET', '../../OOXMLSamples/FormatForMarkup.xml', false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });

    //Insert the markup as text
    myOOXMLRequest.open('GET', fileName, false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'text' });

}

function clickHandler() {

    //This function resets the event handlers for the eleven buttons in the task pane to the function that matches the user's radio button selection
    //See more info on this in the file LoadingAndWritingOOXML.html
    clearButtons();

    if ($('#setOOXMLContent').is(':checked')) {

        $('#setFText').on("click", function () { writeContent('../../OOXMLSamples/TextWithDirectFormat.xml'); });
        $('#setSText').on("click", function () { writeContent('../../OOXMLSamples/TextWithStyle.xml'); });
        $('#setImage').on("click", function () { writeContent('../../OOXMLSamples/SimpleImage.xml'); });
        $('#setPhoto').on("click", function () { writeContent('../../OOXMLSamples/FormattedImage.xml'); });
        $('#setBox').on("click", function () { writeContent('../../OOXMLSamples/TextBoxWordArt.xml'); });
        $('#setShape').on("click", function () { writeContent('../../OOXMLSamples/ShapeWithText.xml'); });
        $('#setControl').on("click", function () { writeContent('../../OOXMLSamples/ContentControl.xml'); });
        $('#setFTable').on("click", function () { writeContent('../../OOXMLSamples/TableWithDirectFormat.xml'); });2
        $('#setSTable').on("click", function () { writeContent('../../OOXMLSamples/TableStyled.xml'); });
        $('#setSmartArt').on("click", function () { writeContent('../../OOXMLSamples/SmartArt.xml'); });
        $('#setChart').on("click", function () { writeContent('../../OOXMLSamples/Chart.xml'); });
    }

    else {
        $('#setFText').on("click", function () { writeMarkup('../../OOXMLSamples/TextWithDirectFormat.xml'); });
        $('#setSText').on("click", function () { writeMarkup('../../OOXMLSamples/TextWithStyle.xml'); });
        $('#setImage').on("click", function () { writeMarkup('../../OOXMLSamples/SimpleImage.xml'); });
        $('#setPhoto').on("click", function () { writeMarkup('../../OOXMLSamples/FormattedImageMarkup.xml'); });
        $('#setBox').on("click", function () { writeMarkup('../../OOXMLSamples/TextBoxWordArt.xml'); });
        $('#setShape').on("click", function () { writeMarkup('../../OOXMLSamples/ShapeWithText.xml'); });
        $('#setControl').on("click", function () { writeMarkup('../../OOXMLSamples/ContentControl.xml'); });
        $('#setFTable').on("click", function () { writeMarkup('../../OOXMLSamples/TableWithDirectFormat.xml'); });
        $('#setSTable').on("click", function () { writeMarkup('../../OOXMLSamples/TableStyled.xml'); });
        $('#setSmartArt').on("click", function () { writeMarkup('../../OOXMLSamples/SmartArt.xml'); });
        $('#setChart').on("click", function () { writeMarkup('../../OOXMLSamples/ChartMarkup.xml'); })
    }
}

function clearButtons() {

    $('#setFText').unbind('click');
    $('#setSText').unbind('click');
    $('#setImage').unbind('click');
    $('#setPhoto').unbind('click');
    $('#setBox').unbind('click');
    $('#setShape').unbind('click');
    $('#setControl').unbind('click');
    $('#setFTable').unbind('click');
    $('#setSTable').unbind('click');
    $('#setSmartArt').unbind('click');
    $('#setChart').unbind('click');

}
// *********************************************************
//
// Word-Add-in-Load-and-write-Open-XML, https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
