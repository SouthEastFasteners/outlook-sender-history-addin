/* global Office */

Office.onReady(() => {
  // Ready for launch-event handlers below.
});

/**
 * Fires when the user opens a message to read, per the LaunchEvent
 * registered in manifest.xml. SendMode="SoftBlock" shows the task pane
 * automatically but still lets the message render without waiting on
 * this handler.
 */
function onMessageRead(event) {
  Office.context.mailbox.item.showTaskpane
    ? Office.context.mailbox.item.showTaskpane()
    : null;
  event.completed();
}

// Register for the manifest's FunctionName reference.
if (typeof window !== "undefined") {
  window.onMessageRead = onMessageRead;
}
