/* function-file.js: Smart Alert onMessageSend implementation */

(function(){
  /**
   * Associate the handler after Office initializes.
   */
  Office.initialize = function(){
    console.log("Office initialized - associating actions.");
    if (Office && Office.actions && Office.actions.associate) {
      Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    } else {
      console.error("Office.actions.associate not available.");
    }
  };

  /**
   * Handler for OnMessageSend LaunchEvent.
   * @param {Office.Event} event
   */
  function onMessageSendHandler(event) {
    try {
      console.log("onMessageSendHandler invoked");
      Office.context.mailbox.item.body.getAsync("text", function(bodyResult){
        if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Failed to get body: ", bodyResult.error);
          // In case of error, allow send so user is not blocked.
          event.completed({ allowEvent: true });
          return;
        }
        var bodyText = bodyResult.value || "";
        var subject = Office.context.mailbox.item.subject || "";
        var contains = containsTestWord(subject, bodyText);
        console.log("Contains 'test':", contains);
        if (!contains) {
          event.completed({ allowEvent: true });
          return;
        }
        // Show confirmation (two logical buttons) using confirm() due to UI-less context.
        // In Smart Alerts capable environments, this will still pause send until event.completed is called.
        var proceed = confirm("Your message contains the word 'test'.\nSend anyway?");
        if (proceed) {
          event.completed({ allowEvent: true });
        } else {
          event.completed({ allowEvent: false });
        }
      });
    } catch(err) {
      console.error("Handler error", err);
      event.completed({ allowEvent: true });
    }
  }

  function containsTestWord(subject, body) {
    var regex = /\btest\b/i;
    return regex.test(subject) || regex.test(body);
  }
})();
