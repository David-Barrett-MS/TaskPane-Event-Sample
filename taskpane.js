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
  console.log(FormatLog("Hooked up button events"));
}

function clearSessionData()
{
  if (Office.context.mailbox.item.sessionData !== undefined) {
    console.log(FormatLog("Office.context.mailbox.item.sessionData is defined"));
      Office.context.mailbox.item.sessionData.clearAsync(function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log(FormatLog("Session data cleared"));
          } else {
              console.log(FormatLog("Failed to clear session data. Error: " + JSON.stringify(asyncResult.error)));
          }
      });
  }
  else
  {
      console.log(FormatLog("Office.context.mailbox.item.sessionData is not defined"));
  }
}

function clearLocalStorage()
{
  if (isLocalStorageAvailable()) {
    localStorage.clear();
    console.log(FormatLog("Local storage cleared"));
  } else {
    console.log(FormatLog("Local storage not available"));
  }
}

function clearStorage() {
    console.log(FormatLog("Clearing all storage"));
    clearLocalStorage();
    clearSessionData();
}

function dumpLocalStorage()
{
    let localData = "Local storage not available";
    if (isLocalStorageAvailable()) {
      localData = JSON.stringify(localStorage, null, 2);
    }
    console.log(FormatLog("Local storage:"));
    console.log(FormatLog(localData));
    document.getElementById("localStorage").value = localData;
}

function dumpSessionData()
{
    console.log(FormatLog("Session data: "));
    if (Office.context.mailbox.item.sessionData !== undefined)
    {
      Office.context.mailbox.item.sessionData.getAllAsync(function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log(FormatLog(JSON.stringify(asyncResult.value)));
              document.getElementById("sessionData").value = JSON.stringify(asyncResult.value, null, 2);
          } else {
              console.log(FormatLog("Failed to get all sessionData. Error: " + JSON.stringify(asyncResult.error)));
              document.getElementById("sessionData").value = "Error: " + JSON.stringify(asyncResult.error);
          }
      });
    }
    else
    {
      console.log(FormatLog("Office.context.mailbox.item.sessionData is not defined"));
      document.getElementById("sessionData").value = "Office.context.mailbox.item.sessionData is not defined";
    }
}

function dumpStorage() {
    console.log(FormatLog("Dumping all storage"));

    dumpLocalStorage();
    dumpSessionData();
}

function setData() {
    setSharedData("SetByTaskPane", new Date().toISOString());
}
