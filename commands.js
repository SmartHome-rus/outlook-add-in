/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SENSITIVE_KEYWORDS: function() { return /* binding */ SENSITIVE_KEYWORDS; },
/* harmony export */   onAppointmentSendHandler: function() { return /* binding */ onAppointmentSendHandler; },
/* harmony export */   onMessageSendHandler: function() { return /* binding */ onMessageSendHandler; },
/* harmony export */   onNewAppointmentOrganizerHandler: function() { return /* binding */ onNewAppointmentOrganizerHandler; },
/* harmony export */   onNewMessageComposeHandler: function() { return /* binding */ onNewMessageComposeHandler; },
/* harmony export */   scanForSensitiveKeywords: function() { return /* binding */ scanForSensitiveKeywords; }
/* harmony export */ });
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/// <reference types="office-js" />
// Office.js initialization with enhanced logging
Office.onReady((info) => {
    var _a, _b, _c;
    console.log('✅ Office.js is ready');
    console.log('📊 Host:', info.host);
    console.log('📊 Platform:', info.platform);
    console.log('📊 Context available:', !!Office.context);
    console.log('📊 Mailbox available:', !!((_a = Office.context) === null || _a === void 0 ? void 0 : _a.mailbox));
    console.log('📊 Item available:', !!((_c = (_b = Office.context) === null || _b === void 0 ? void 0 : _b.mailbox) === null || _c === void 0 ? void 0 : _c.item));
    // Verify event handlers are properly registered
    if (typeof window !== 'undefined') {
        console.log('📦 Event handlers registered:', {
            onMessageSendHandler: typeof window.onMessageSendHandler,
            onAppointmentSendHandler: typeof window.onAppointmentSendHandler
        });
    }
});
/**
 * Predefined list of sensitive keywords to scan for.
 * This list is hardcoded and can be modified as needed.
 */
const SENSITIVE_KEYWORDS = [
    'confidential',
    'internal use only',
    'secret',
    'project-alpha',
    'classified',
    'proprietary',
    'restricted',
    'private',
    'sensitive',
    'do not distribute',
    'for internal use',
    'confidential information',
    'trade secret',
    'intellectual property',
    'proprietary information'
];
/**
 * Scans the provided text for sensitive keywords.
 * @param text The text content to scan
 * @returns An object containing found keywords and whether any were detected
 */
function scanForSensitiveKeywords(text) {
    if (!text) {
        return { hasKeywords: false, foundKeywords: [] };
    }
    const foundKeywords = [];
    const lowerCaseText = text.toLowerCase();
    for (const keyword of SENSITIVE_KEYWORDS) {
        if (lowerCaseText.includes(keyword.toLowerCase())) {
            foundKeywords.push(keyword);
        }
    }
    return {
        hasKeywords: foundKeywords.length > 0,
        foundKeywords: foundKeywords
    };
}
/**
 * Gets the body content of the current item (email or appointment).
 * @param context The Office context
 * @returns Promise resolving to the body text
 */
async function getItemBody(context) {
    return new Promise((resolve, reject) => {
        context.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
            var _a;
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                resolve(asyncResult.value || '');
            }
            else {
                reject(new Error(`Failed to get body: ${(_a = asyncResult.error) === null || _a === void 0 ? void 0 : _a.message}`));
            }
        });
    });
}
/**
 * Handler for the OnMessageSend event.
 * This function is called when a user attempts to send an email.
 */
function onMessageSendHandler(event) {
    var _a, _b;
    console.log('OnMessageSend event triggered');
    // Set up timeout protection (3 seconds max)
    const timeoutId = setTimeout(() => {
        console.warn('⚠️ OnMessageSend handler timed out - allowing send to proceed');
        event.completed({ allowEvent: true });
    }, 3000);
    try {
        // Check if Office is ready and item exists
        if (!((_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.item)) {
            console.warn('⚠️ Office context or item not available - allowing send');
            clearTimeout(timeoutId);
            event.completed({ allowEvent: true });
            return;
        }
        // Quick synchronous check first (if possible)
        const item = Office.context.mailbox.item;
        if (!item.body) {
            console.warn('⚠️ Item body not accessible - allowing send');
            clearTimeout(timeoutId);
            event.completed({ allowEvent: true });
            return;
        }
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
            var _a;
            clearTimeout(timeoutId); // Clear timeout since we got a response
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const bodyText = asyncResult.value || '';
                const scanResult = scanForSensitiveKeywords(bodyText);
                if (scanResult.hasKeywords) {
                    console.log('🚫 Sensitive keywords found:', scanResult.foundKeywords);
                    // Block the send and show error message
                    const errorMessage = `This email contains potentially sensitive content (found: ${scanResult.foundKeywords.join(', ')}). Please review before sending.`;
                    event.completed({
                        allowEvent: false,
                        errorMessage: errorMessage
                    });
                }
                else {
                    console.log('✅ No sensitive keywords found, allowing send');
                    // Allow the send to proceed
                    event.completed({ allowEvent: true });
                }
            }
            else {
                console.error('❌ Failed to get email body:', (_a = asyncResult.error) === null || _a === void 0 ? void 0 : _a.message);
                // In case of error getting body, allow send to proceed (fail-safe)
                event.completed({ allowEvent: true });
            }
        });
    }
    catch (error) {
        console.error('❌ Exception in onMessageSendHandler:', error);
        clearTimeout(timeoutId);
        // In case of any exception, allow send to proceed (fail-safe)
        event.completed({ allowEvent: true });
    }
}
/**
 * Handler for the OnAppointmentSend event.
 * This function is called when a user attempts to send a calendar appointment.
 */
function onAppointmentSendHandler(event) {
    var _a, _b;
    console.log('OnAppointmentSend event triggered');
    // Set up timeout protection (3 seconds max)
    const timeoutId = setTimeout(() => {
        console.warn('⚠️ OnAppointmentSend handler timed out - allowing send to proceed');
        event.completed({ allowEvent: true });
    }, 3000);
    try {
        // Check if Office is ready and item exists
        if (!((_b = (_a = Office === null || Office === void 0 ? void 0 : Office.context) === null || _a === void 0 ? void 0 : _a.mailbox) === null || _b === void 0 ? void 0 : _b.item)) {
            console.warn('⚠️ Office context or item not available - allowing appointment send');
            clearTimeout(timeoutId);
            event.completed({ allowEvent: true });
            return;
        }
        // Quick synchronous check first (if possible)
        const item = Office.context.mailbox.item;
        if (!item.body) {
            console.warn('⚠️ Item body not accessible - allowing appointment send');
            clearTimeout(timeoutId);
            event.completed({ allowEvent: true });
            return;
        }
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
            var _a;
            clearTimeout(timeoutId); // Clear timeout since we got a response
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const bodyText = asyncResult.value || '';
                const scanResult = scanForSensitiveKeywords(bodyText);
                if (scanResult.hasKeywords) {
                    console.log('🚫 Sensitive keywords found in appointment:', scanResult.foundKeywords);
                    // Block the send and show error message
                    const errorMessage = `This appointment contains potentially sensitive content (found: ${scanResult.foundKeywords.join(', ')}). Please review before sending.`;
                    event.completed({
                        allowEvent: false,
                        errorMessage: errorMessage
                    });
                }
                else {
                    console.log('✅ No sensitive keywords found in appointment, allowing send');
                    // Allow the send to proceed
                    event.completed({ allowEvent: true });
                }
            }
            else {
                console.error('❌ Failed to get appointment body:', (_a = asyncResult.error) === null || _a === void 0 ? void 0 : _a.message);
                // In case of error getting body, allow send to proceed (fail-safe)
                event.completed({ allowEvent: true });
            }
        });
    }
    catch (error) {
        console.error('❌ Exception in onAppointmentSendHandler:', error);
        clearTimeout(timeoutId);
        // In case of any exception, allow send to proceed (fail-safe)
        event.completed({ allowEvent: true });
    }
}
/**
 * Handler for the OnNewMessageCompose event.
 * This function is called when a new email composition is started.
 */
function onNewMessageComposeHandler(event) {
    console.log('OnNewMessageCompose event triggered');
    // No specific action needed for compose start
    event.completed();
}
/**
 * Handler for the OnNewAppointmentOrganizer event.
 * This function is called when a new appointment is being organized.
 */
function onNewAppointmentOrganizerHandler(event) {
    console.log('OnNewAppointmentOrganizer event triggered');
    // No specific action needed for appointment organizer start
    event.completed();
}
// Ensure the functions are available in the global scope for Office.js to call them
if (typeof window !== 'undefined') {
    window.onMessageSendHandler = onMessageSendHandler;
    window.onAppointmentSendHandler = onAppointmentSendHandler;
    window.onNewMessageComposeHandler = onNewMessageComposeHandler;
    window.onNewAppointmentOrganizerHandler = onNewAppointmentOrganizerHandler;
}
// For Node.js environments (testing)
if (typeof __webpack_require__.g !== 'undefined') {
    __webpack_require__.g.onMessageSendHandler = onMessageSendHandler;
    __webpack_require__.g.onAppointmentSendHandler = onAppointmentSendHandler;
    __webpack_require__.g.onNewMessageComposeHandler = onNewMessageComposeHandler;
    __webpack_require__.g.onNewAppointmentOrganizerHandler = onNewAppointmentOrganizerHandler;
}
// Export functions for testing


/******/ })()
;
//# sourceMappingURL=commands.js.map