function onMessageSendHandler(event) {
  // Add message validation or processing here.
  event.completed({ allowEvent: true });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);