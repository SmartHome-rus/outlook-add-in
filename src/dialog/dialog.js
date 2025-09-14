/**
 * Test Word Checker Dialog
 * Handles user interaction in the confirmation dialog
 */

/**
 * Office.js initialization
 */
Office.onReady((info) => {
    console.log("Dialog initialized");
    
    if (info.host === Office.HostType.Outlook) {
        initializeDialog();
    }
});

/**
 * Initialize the dialog functionality
 */
function initializeDialog() {
    try {
        // Get button elements
        const sendAnywayBtn = document.getElementById('send-anyway');
        const discardEditBtn = document.getElementById('discard-edit');
        
        if (!sendAnywayBtn || !discardEditBtn) {
            console.error("Could not find dialog buttons");
            return;
        }
        
        // Add click event listeners
        sendAnywayBtn.addEventListener('click', handleSendAnyway);
        discardEditBtn.addEventListener('click', handleDiscardEdit);
        
        // Add keyboard event listeners for accessibility
        document.addEventListener('keydown', handleKeyPress);
        
        // Focus the first button for better accessibility
        sendAnywayBtn.focus();
        
        console.log("Dialog event listeners initialized");
        
    } catch (error) {
        console.error("Error initializing dialog:", error);
    }
}

/**
 * Handle "Send Anyway" button click
 */
function handleSendAnyway() {
    console.log("User clicked Send Anyway");
    
    try {
        // Disable buttons to prevent multiple clicks
        disableButtons();
        
        // Send message to parent (commands.js) indicating the user wants to send
        sendMessageToParent("send");
        
    } catch (error) {
        console.error("Error in handleSendAnyway:", error);
        
        // Re-enable buttons if there's an error
        enableButtons();
    }
}

/**
 * Handle "Discard & Edit" button click
 */
function handleDiscardEdit() {
    console.log("User clicked Discard & Edit");
    
    try {
        // Disable buttons to prevent multiple clicks
        disableButtons();
        
        // Send message to parent (commands.js) indicating the user wants to cancel
        sendMessageToParent("discard");
        
    } catch (error) {
        console.error("Error in handleDiscardEdit:", error);
        
        // Re-enable buttons if there's an error
        enableButtons();
    }
}

/**
 * Handle keyboard events for better accessibility
 * @param {KeyboardEvent} event - The keyboard event
 */
function handleKeyPress(event) {
    if (event.key === 'Escape') {
        // ESC key should cancel (same as Discard & Edit)
        handleDiscardEdit();
    } else if (event.key === 'Enter') {
        // Enter key should focus the primary action
        const activeElement = document.activeElement;
        if (activeElement && activeElement.id) {
            activeElement.click();
        } else {
            // Default to "Send Anyway" if no specific focus
            handleSendAnyway();
        }
    }
}

/**
 * Send a message to the parent window (commands.js)
 * @param {string} message - The message to send ("send" or "discard")
 */
function sendMessageToParent(message) {
    try {
        console.log("Sending message to parent:", message);
        
        // Primary method: Use Office.context.ui.messageParent
        if (Office && Office.context && Office.context.ui && Office.context.ui.messageParent) {
            Office.context.ui.messageParent(message);
            
            // Close the dialog after sending the message
            setTimeout(() => {
                closeDialog();
            }, 100);
            
        } else {
            console.error("Office.context.ui.messageParent is not available");
            
            // Fallback: Try to call parent function directly (for testing)
            if (window.parent && window.parent.receiveDialogMessage) {
                window.parent.receiveDialogMessage(message);
                closeDialog();
            } else {
                console.error("No communication method available with parent");
            }
        }
        
    } catch (error) {
        console.error("Error sending message to parent:", error);
        
        // If all else fails, at least close the dialog
        closeDialog();
    }
}

/**
 * Close the dialog
 */
function closeDialog() {
    try {
        if (Office && Office.context && Office.context.ui && Office.context.ui.closeContainer) {
            Office.context.ui.closeContainer();
        } else {
            console.error("Office.context.ui.closeContainer is not available");
            
            // Fallback: Try to close the window
            if (window.close) {
                window.close();
            }
        }
    } catch (error) {
        console.error("Error closing dialog:", error);
    }
}

/**
 * Disable dialog buttons to prevent multiple clicks
 */
function disableButtons() {
    const sendAnywayBtn = document.getElementById('send-anyway');
    const discardEditBtn = document.getElementById('discard-edit');
    
    if (sendAnywayBtn) {
        sendAnywayBtn.disabled = true;
        sendAnywayBtn.textContent = 'Sending...';
    }
    
    if (discardEditBtn) {
        discardEditBtn.disabled = true;
    }
}

/**
 * Re-enable dialog buttons (in case of error)
 */
function enableButtons() {
    const sendAnywayBtn = document.getElementById('send-anyway');
    const discardEditBtn = document.getElementById('discard-edit');
    
    if (sendAnywayBtn) {
        sendAnywayBtn.disabled = false;
        sendAnywayBtn.textContent = 'Send Anyway';
    }
    
    if (discardEditBtn) {
        discardEditBtn.disabled = false;
    }
}

/**
 * Alternative function that can be called directly from the parent
 * (for compatibility with different Office.js versions)
 */
function notifyParent(action) {
    sendMessageToParent(action);
}

// Make functions globally available for debugging and alternative communication
if (typeof window !== "undefined") {
    window.handleSendAnyway = handleSendAnyway;
    window.handleDiscardEdit = handleDiscardEdit;
    window.notifyParent = notifyParent;
    window.closeDialog = closeDialog;
}

// For Node.js environments (if needed for testing)
if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        handleSendAnyway,
        handleDiscardEdit,
        notifyParent,
        closeDialog
    };
}