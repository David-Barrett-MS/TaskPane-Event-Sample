# Sharing data between an event and a TaskPane

This sample sets localStorage and Office sessionData from the event of an Outlook add-in (messageSend, appointmentSend, messageRecipientsChanged) and implements a taskpane that allows displaying of the same data.  This allows testing to see which storage mechanisms are available in specific scenarios.

The sample uses only office.js calls, so needs no application registration.  It can be installed directly from Github using the [Github manifest](https://raw.githubusercontent.com/David-Barrett-MS/TaskPane-Event-Sample/refs/heads/main/TaskPaneEventSample%20Github.xml) (the add-in files are served from Github Pages deployed from this repository).

## Testing

The add-in activates on message or appointment compose.  From the compose screen:

- Open the add-in TaskPane.
- Click Dump storage data to see what data is currently held in the various storages.
- Click Clear all storage, then Dump storage data again (and confirm that both Local Storage and Office Session Data are now empty).
- Add a recipient (or attendee) to the item.  This will trigger an event.  When the event has finished processing, you'll see a notification at the top of the compose window showing OnMessageRecipientsChangedCalled.  The event attempts to write a variable (with the current time) to both local storage and Office session data.
- Click Dump storage data.  You'll see the variable that was set during the event in one or both of the storages (local or office session).

## Test Results

In my tests, both local storage and Office session data could be used to share data between the event and the TaskPane in all Outlook desktop and web clients (mobile clients were not tested).  Tested clients: Outlook for Mac, OWA in Edge on Mac, OWA in Safari on Mac, Outlook Classic, New Outlook, OWA on Windows Edge.