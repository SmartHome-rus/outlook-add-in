/*
 * Outlook Smart Alert Add-in
 * 
 * This add-in uses Smart Alerts to warn users when sending emails containing the word "test".
 * It implements a two-phase send process:
 * 1. Initial send attempt: Block and show notification if "test" is found
 * 2. Second send attempt: Allow send after user confirms via "Send Anyway" button
 */

// Global flag key for session data
const SEND_APPROVED_KEY = "isSendApproved";
const NOTIFICATION_KEY = "test-alert";

/**
 * Main handler for OnMessageSend event
 * Implements Smart Alerts workflow for detecting "test" keyword
 */
async function onMessageSendHandler(event) {
    try {
        console.log("OnMessageSend event triggered");
        
        // PHASE 1: Check if this is a second send attempt (user already approved)
        // If user clicked "Send Anyway", we set a flag in session data
        const item = Office.context.mailbox.item;
        
        // Check for approval flag from previous interaction
        const sendApproved = await checkSendApproval();
        if (sendApproved) {
            console.log("Send already approved by user, allowing email to send");
            // Clear the approval flag and allow send
            await clearSendApproval();
            event.completed({ allowEvent: true });
            return;
        }

        // PHASE 2: Offline check - if offline, skip all logic and allow send
        if (!navigator.onLine) {
            console.log("User is offline, allowing email to go to Outbox");
            event.completed({ allowEvent: true });
            return;
        }

        // PHASE 3: Get email body and scan for "test" keyword
        const bodyText = await getEmailBodyText();
        const containsTest = /\btest\b/i.test(bodyText);

        if (!containsTest) {
            console.log("No 'test' keyword found, allowing send");
            event.completed({ allowEvent: true });
            return;
        }

        // PHASE 4: "test" keyword found - show Smart Alert and block send
        console.log("'test' keyword detected, showing Smart Alert");
        await showTestWarningNotification();
        
        // Block the send - user must click "Send Anyway" to proceed
        event.completed({ allowEvent: false });

    } catch (error) {
        console.error("Error in onMessageSendHandler:", error);
        // On error, allow send to avoid blocking legitimate emails
        event.completed({ allowEvent: true });
    }
}

/**
 * Action handler for "Send Anyway" button click
 * This function is called when user clicks the Smart Alert action button
 */
async function sendAnywayAction(event) {
    try {
        console.log("Send Anyway button clicked");
        
        // Set approval flag in session data for next send attempt
        await setSendApproval();
        
        // Remove the notification bar
        await removeTestWarningNotification();
        
        console.log("Notification removed, user can now send email");
        
        // Complete the action
        event.completed();
        
    } catch (error) {
        console.error("Error in sendAnywayAction:", error);
        event.completed();
    }
}

/**
 * Get the email body as plain text
 */
async function getEmailBodyText() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value || "");
            } else {
                reject(new Error("Failed to get email body: " + result.error.message));
            }
        });
    });
}

/**
 * Show Smart Alert notification warning about "test" keyword
 */
async function showTestWarningNotification() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        const notification = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "This email contains the word 'test'. Are you sure you want to send it?",
            icon: "Icon.80x80",
            persistent: false,
            actions: [
                {
                    actionType: Office.MailboxEnums.ActionType.ShowTaskpane,
                    actionText: "Send Anyway",
                    commandId: "sendAnywayAction"
                }
            ]
        };

        item.notificationMessages.addAsync(NOTIFICATION_KEY, notification, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Smart Alert notification added successfully");
                resolve();
            } else {
                reject(new Error("Failed to add notification: " + result.error.message));
            }
        });
    });
}

/**
 * Remove the test warning notification
 */
async function removeTestWarningNotification() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        item.notificationMessages.removeAsync(NOTIFICATION_KEY, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Notification removed successfully");
                resolve();
            } else {
                console.warn("Failed to remove notification:", result.error.message);
                // Don't reject - this is not critical
                resolve();
            }
        });
    });
}

/**
 * Set the send approval flag in session data
 */
async function setSendApproval() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        item.sessionData.setAsync(SEND_APPROVED_KEY, "true", (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Send approval flag set");
                resolve();
            } else {
                reject(new Error("Failed to set approval flag: " + result.error.message));
            }
        });
    });
}

/**
 * Check if send has been approved (flag exists in session data)
 */
async function checkSendApproval() {
    return new Promise((resolve) => {
        const item = Office.context.mailbox.item;
        
        item.sessionData.getAsync(SEND_APPROVED_KEY, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value === "true") {
                resolve(true);
            } else {
                resolve(false);
            }
        });
    });
}

/**
 * Clear the send approval flag from session data
 */
async function clearSendApproval() {
    return new Promise((resolve) => {
        const item = Office.context.mailbox.item;
        
        item.sessionData.removeAsync(SEND_APPROVED_KEY, (result) => {
            // Always resolve - clearing is not critical for functionality
            resolve();
        });
    });
}

// Associate the action command with its handler function
// This must be done at the global level when the script loads
Office.onReady(() => {
    console.log("Office.js is ready, associating action handlers");
    
    // Associate the "Send Anyway" button action with its handler
    Office.actions.associate("sendAnywayAction", sendAnywayAction);
    
    console.log("Smart Alert add-in loaded successfully");
});