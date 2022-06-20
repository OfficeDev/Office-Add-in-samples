/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to display a message on the task pane.
 */

function showMessage(text) {
   const appendedText = $('#message-area').html() + text + "<br>---";
    $('.welcome-body').hide();
    $('#message-area').show(); 
    $('#message-area').html(appendedText);
 }

 function clearMessage() {
   $('.welcome-body').hide();
   $('#message-area').show(); 
   $('#message-area').html("---<br>");
 }