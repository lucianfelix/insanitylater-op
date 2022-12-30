/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action!",
    icon: "Icon.80x80",
    persistent: false,
  };

  console.log("test");

  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  //get the current message body
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Text,
    { asyncContext: "This is passed to the callback" },
    function (asyncResult: Office.AsyncResult<string>) {
      const body = asyncResult.value;

      const newBody = body + "This is a new body";

      Office.context.mailbox.item.body.setAsync(
        newBody,
        { asyncContext: "This is passed to the callback" },
        function () {
          // // Show a notification message
          Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
          // // Office.context.mailbox.item.saveAsync();
          // //console.log(result);
        }
      );
      // console.log(result.value);
      // console.log(result.asyncContext);
    }
  );

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
