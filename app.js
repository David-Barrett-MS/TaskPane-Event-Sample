/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var statusInfo = "";
var fullLogEventAPIUrl = ""; // The API URL including any additional parameters
var baseLogEventAPIUrl = ""; // The API URL
var addinSettings;
const AddinName = "TaskPaneEventSample";

//Office.onReady();
Office.initialize = function () {
    // This function is not called during OnMessageSend LaunchEvent in Outlook Desktop, so any initialisation here won't work in that scenario
}

function FormatLog(data) {
    // Return log with add-in name and current time prepended
  let now = new Date();
  let currentTime = now.toLocaleTimeString('en-US', { hour12: false });
  return AddinName + " " + currentTime + ": " + data;
}

async function getInsightMessage() {
  return {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: "This is an InsightMessage",
    icon: "Icon.16x16",
    actions: [
      {
        actionText: "Process manually",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
        commandId: "msgComposeOpenPaneButton",
        contextData: "{}"
      }
    ]
  };
}

async function applyInsightMessage(event) {
  const notification = await getInsightMessage();

  console.log(FormatLog("Applying InsightMessage:"), notification);
  Office.context.mailbox.item.notificationMessages.replaceAsync("InsightMessage", notification, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to apply InsightMessage:", asyncResult.error.message);
      return;
    }
    console.log(FormatLog("InsightMessage applied"));
  });

  if (event) {
    event.completed();
  }
}

/**
 * Set notification on MailItem (overwrites any previous notification)
 * @param {Notification message to be set} message 
 */
async function SetNotification(message) {
    let infoMessage =
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "icon2",
      persistent: true
    };    
    Office.context.mailbox.item.notificationMessages.replaceAsync(AddinName + "Notification", infoMessage);
}

/**
 * Append the given status to the notification for the MailItem
 * @param {Message to be added to the status} message 
 * @returns 
 */
async function SetStatus(message) {
    if (statusInfo != "") {
        statusInfo = statusInfo + " | ";    
    }
    statusInfo = statusInfo + message;
    console.log(FormatLog(message));
    return SetNotification(statusInfo);
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

function setInLocalStorage(key, value) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned. 
  // If so, use the partition to ensure the data is only accessible by your add-in.
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}

function getFromLocalStorage(key) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}

function setSharedData(key, value) {
    // Set data that should be accessible from the TaskPane

    if (isLocalStorageAvailable()) {
        setInLocalStorage(key, value);
        console.log(FormatLog("Set local storage: " + key + " = " + value));
    }
    else {
        console.log(FormatLog("Local storage not available"));
    }
    Office.context.mailbox.item.sessionData.setAsync(key, value);
    console.log(FormatLog("Set session data: " + key + " = " + value));
}

function getSharedData(key) {
    // Get data that should be accessible from the TaskPane
    let sessionDataAvailable = true;

    if (Office.context.platform === Office.PlatformType.Mac) { //} && !                navigator.userAgent.includes("OneOutlook")) {
        // Mac Outlook
        sessionDataAvailable = false;
        console.log(FormatLog("Reading " + key + "from local storage"));
        return getFromLocalStorage(key);
    }

    if (sessionDataAvailable) {
        console.log(FormatLog("Reading " + key + "from session data"));
        Office.context.mailbox.item.sessionData.getAsync(key, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                return asyncResult.value;
            }
            else {
                return null;
            }
        });
    }
}

function dumpStorageByKey(key) {
    console.log(FormatLog("Dumping storage for key: " + key));
    let value = getFromLocalStorage(key);
    console.log(FormatLog("Local storage: " + key + " = " + value));

    Office.context.mailbox.item.sessionData.getAsync(key, function (asyncResult) {
        console.log(FormatLog("Session data: " + key + " = " + asyncResult.value));
    });
}

function onMessageSendHandler(event) {
    setSharedData("onMessageSendCalled", new Date().toISOString());
    SetStatus("onMessageSendCalled");
    event.completed({ allowEvent: false });
}

function OnAppointmentSendHandler(event) {
    setSharedData("onAppointmentSendCalled", new Date().toISOString());
    SetStatus("onAppointmentSendCalled");
    event.completed({ allowEvent: false });
}

function OnMessageRecipientsChangedHandler(event) {
    setSharedData("OnMessageRecipientsChangedCalled", new Date().toISOString());
    SetStatus("OnMessageRecipientsChangedCalled");
    event.completed({ allowEvent: true });
}

function OnAppointmentAttendeesChangedHandler(event) {
    setSharedData("OnAppointmentAttendeesChangedCalled", new Date().toISOString());
    SetStatus("OnAppointmentAttendeesChangedCalled");
    event.completed({ allowEvent: true });
}


if (Office.context !== undefined && (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) ) {
    // Associate the events with their respective handlers
    Office.actions.associate("OnMessageSendHandler", onMessageSendHandler);
    Office.actions.associate("OnAppointmentSendHandler", OnAppointmentSendHandler);
    Office.actions.associate("OnMessageRecipientsChangedHandler", OnMessageRecipientsChangedHandler);
    Office.actions.associate("OnAppointmentAttendeesChangedHandler", OnAppointmentAttendeesChangedHandler);
}