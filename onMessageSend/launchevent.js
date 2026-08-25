console.log("onMessageSend/launchevent.js loaded");

Office.initialize = function () {
}

function onMessageSendHandler(event) {
  // Add message validation or processing here.
  console.log("onMessageSendHandler called");
  // TODO
  event.completed({ allowEvent: true });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);