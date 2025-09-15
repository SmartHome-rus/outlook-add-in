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

        // Soft-block pattern: first time we see the word 'test', show notification and cancel send.
        // If user presses send again (we set a custom flag), allow it.
        var item = Office.context.mailbox.item;
        var customPropKey = "smartAlertTestConfirmed";

        // Try to use customProps (async) - but custom properties load cost may exceed timeout.
        // Instead stash a transient flag on the item via session storage keyed by itemId if available.
        var itemId = item.itemId || (item.saveAsync ? "temp-item" : "unknown");
        var cacheKey = "smartAlertSeen_" + itemId;
        var alreadyConfirmed = false;
        try { alreadyConfirmed = window.sessionStorage.getItem(cacheKey) === '1'; } catch(e) { /* ignore */ }

        if (alreadyConfirmed) {
          console.log("Second attempt - allowing send despite 'test'.");
          event.completed({ allowEvent: true });
          return;
        }

        // Show an informational message (Smart Alert style) and block this send.
        try {
          item.notificationMessages.replaceAsync("smartAlertTestWord", {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Message contains the word 'test'. Edit it or press Send again to proceed.",
            icon: "", // no icon
            persistent: true
          }, function(){
            console.log("Notification inserted - soft-blocking first send.");
            try { window.sessionStorage.setItem(cacheKey, '1'); } catch(e) {}
            event.completed({ allowEvent: false });
          });
        } catch(notificationError) {
          console.warn("Notification failed, falling back to soft-block.", notificationError);
          try { window.sessionStorage.setItem(cacheKey, '1'); } catch(e) {}
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
