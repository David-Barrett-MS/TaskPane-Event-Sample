# Sharing data between an event and a TaskPane

This sample sets localStorage and Office sessionData from the event of an Outlook add-in (messageSend, appointmentSend, messageRecipientsChanged) and implements a taskpane that allows displaying of the same data.  This allows testing to see which storage mechanisms are available in specific scenarios.

The sample uses only office.js calls, so needs no application registration.  It can be installed directly from Github using the [Github manifest](https://raw.githubusercontent.com/David-Barrett-MS/TaskPane-Event-Sample/refs/heads/main/TaskPaneEventSample%20Github.xml) (the add-in files are served via Github Pages from this repository).
