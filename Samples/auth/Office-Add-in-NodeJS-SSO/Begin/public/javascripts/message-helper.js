// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file shows how to display a message on the task pane.

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