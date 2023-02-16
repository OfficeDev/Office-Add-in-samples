// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file shows how to display a message on the task pane.

 function showMessage(text) {
  const appendedText = document.getElementById('message-area').innerHTML + text + "<br>---";
  document.getElementById('message-area').innerHTML = appendedText;
 }

 function clearMessage() {
  document.getElementById('welcome-body').style.display = 'none';
  document.getElementById('message-area').style.display = 'display';
  document.getElementById('message-area').innerHTML = "---<br>";
 }