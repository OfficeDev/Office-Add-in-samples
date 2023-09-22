/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { base64Image } from "../../base64Image";
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("insert-image").onclick = () => clearMessage(insertImage);
    document.getElementById("insert-text").onclick = () => clearMessage(insertText);
    document.getElementById("get-slide-metadata").onclick = () => clearMessage(getSlideMetadata);
    document.getElementById("add-slides").onclick = () => tryCatch(addSlides);
    document.getElementById("go-to-first-slide").onclick = () => clearMessage(goToFirstSlide);
    document.getElementById("go-to-next-slide").onclick = () => clearMessage(goToNextSlide);
    document.getElementById("go-to-previous-slide").onclick = () => clearMessage(goToPreviousSlide);
    document.getElementById("go-to-last-slide").onclick = () => clearMessage(goToLastSlide);
  }
});

function insertImage() {
  // Call Office.js to insert the image into the document.
  Office.context.document.setSelectedDataAsync(base64Image, {
      coercionType: Office.CoercionType.Image
  },
      function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            setMessage("Error: " + asyncResult.error.message);
          }
      });
}

function insertText() {
  Office.context.document.setSelectedDataAsync('Hello World!',
      function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            setMessage("Error: " + asyncResult.error.message);
          }
      });
}

function getSlideMetadata() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                setMessage("Error: " + asyncResult.error.message);
            } else {
                setMessage("Metadata for selected slides: " + JSON.stringify(asyncResult.value));
            }
        }
    );
  }
  
  async function addSlides() {
  await PowerPoint.run(async function (context) {
      context.presentation.slides.add();
      context.presentation.slides.add();

      await context.sync();

      goToLastSlide();
      setMessage("Success: Slides added.");
  });
}

function goToFirstSlide() {
  Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
      function (asyncResult) {
          if (asyncResult.status == "failed") {
            setMessage("Error: " + asyncResult.error.message);
          }
      });
}

function goToLastSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
      function (asyncResult) {
          if (asyncResult.status == "failed") {
            setMessage("Error: " + asyncResult.error.message);
          }
      });
}

function goToPreviousSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
      function (asyncResult) {
          if (asyncResult.status == "failed") {
            setMessage("Error: " + asyncResult.error.message);
          }
      });
}

function goToNextSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
      function (asyncResult) {
          if (asyncResult.status == "failed") {
            setMessage("Error: " + asyncResult.error.message);
          }
      });
}

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

function setMessage(message) {
    document.getElementById("message").innerText = message;
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        document.getElementById("message").innerText = "";
        await callback();
    } catch (error) {
        setMessage("Error: " + error.toString());
    }
}
