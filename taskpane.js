/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



/**
 * The Office.initialize function that gets called when the Office.js library is loaded.
 */
Office.initialize = function () {

  // Initialize instance variables to access API objects.

  document.getElementById("dumpStorage").onclick = dumpStorage; // Add the click event for the button
  document.getElementById("setValue").onclick = setData;
  document.getElementById("clearStorage").onclick = clearStorage;
  console.log("Hooked up button events");

  // Set up the ItemChanged event.
  if (Office.context.mailbox.item == null) {
    console.log("Item is null");
  }
}

function clearStorage() {
    console.log("Clearing all storage");
    if (isLocalStorageAvailable()) {
      localStorage.clear();
      console.log("Local storage cleared");
    } else {
      console.log("Local storage not available");
    }
    Office.context.mailbox.item.sessionData.clearAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Session data cleared");
        } else {
            console.log("Failed to clear session data. Error: " + JSON.stringify(asyncResult.error));
        }
    });
}

function dumpStorage() {
    console.log("Dumping all storage");

    console.log("Local storage: ");
    let localData = "Local storage not available";
    if (isLocalStorageAvailable()) {
      localData = JSON.stringify(localStorage, null, 2);
    }
    console.log(localData);
    document.getElementById("localStorage").value = localData;

    console.log("Session storage: ");
    let sessionData = "Session storage not available";
    if (typeof(sessionStorage) !== "undefined") {
      sessionData = JSON.stringify(sessionStorage, null, 2);
    }
    console.log(sessionData);
    document.getElementById("sessionStorage").value = sessionData;

    console.log("Session data: ");
    Office.context.mailbox.item.sessionData.getAllAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(JSON.stringify(asyncResult.value));
            document.getElementById("sessionData").value = JSON.stringify(asyncResult.value, null, 2);
        } else {
            console.log("Failed to get all sessionData. Error: " + JSON.stringify(asyncResult.error));
            document.getElementById("sessionData").value = "Error: " + JSON.stringify(asyncResult.error);
        }
    });
}

function isLocalStorageAvailable() {
  try {
    localStorage.setItem('test', 'test');
    localStorage.removeItem('test');
    return true;
  } catch (e) {
    return false;
  }
}

function setData() {
  let key = "SetByTaskPane";
  let value = new Date().toISOString();
  if (isLocalStorageAvailable()) {
    localStorage.setItem(key, value);
    console.log("Set local storage: " + key + " = " + value);
  }
  Office.context.mailbox.item.sessionData.setAsync(key, value);
  console.log("Set session data: " + key + " = " + value);
}
