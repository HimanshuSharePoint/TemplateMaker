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
 * @param event {Office.AddinCommands.Event}
 */

 async function writeValue(event) {
  Word.run(async (context) => {
    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(
      "ExecuteFunction works. Button ID=" + event.source.id,
      Word.InsertLocation.end
    );

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}


function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

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

const g = getGlobal();

// The add-in command functions need to be available in global scope
//g.action = action;
g.writeValue = writeValue;
