/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to display a message on the task pane.
 */

 function showMessage(text) {
    $('.welcome-body').hide();
    $('#message-area').show(); 
    $('#message-area').text(text);
 }