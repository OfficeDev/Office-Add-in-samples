/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
*/

/* global document, Office */
var myLocalStorage;

Office.onReady((info) => {
  console.log(`Office.onReady(): Host: ${Office.HostType.Outlook}`)
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";    

    //Just testing to make sure localStorage is working in Outlook Desktop. Works in a Task Pane, but not in commands.js!
    // if (myLocalStorage === undefined){
    //   console.log("Office.onReady(): localStorage var not set: setting...");  
    //   myLocalStorage = window.localStorage;

    //   try {        
    //     myLocalStorage.setItem("foo", "FOOOOOO DATA!");
    //     console.log("Office.onReady()(): myLocalStorage set");
    //   }
    //   catch(ex){
    //     console.error(`Office.onReady()(): Error calling localStorage: ${ex}`);
    //   }  
    // }    
  }
});

function run() {
  /**
   * Insert your Outlook code here
   */
}
