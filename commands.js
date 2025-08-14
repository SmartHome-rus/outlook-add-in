/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   onAppointmentSendHandler: () => (/* binding */ onAppointmentSendHandler),
/* harmony export */   onMessageSendHandler: () => (/* binding */ onMessageSendHandler),
/* harmony export */   scanForSensitiveKeywords: () => (/* binding */ scanForSensitiveKeywords)
/* harmony export */ });
/**
 * Sensitive Data Scanner - On-Send Event Handlers
 * This file contains the core logic for scanning emails and appointments
 * for sensitive keywords and blocking sending when found.
 */
/// <reference path="../types/office.d.ts" />
// Define the list of sensitive keywords (case-insensitive)
const SENSITIVE_KEYWORDS = ['confidential', 'internal use only', 'secret', 'project-alpha', 'classified', 'restricted', 'proprietary', 'do not distribute', 'sensitive', 'private and confidential'];
/**
 * Scans text content for sensitive keywords
 * @param content - The text content to scan
 * @returns Array of found keywords
 */
function scanForSensitiveKeywords(content) {
  if (!content) return [];
  const normalizedContent = content.toLowerCase();
  const foundKeywords = [];
  for (const keyword of SENSITIVE_KEYWORDS) {
    if (normalizedContent.includes(keyword.toLowerCase())) {
      foundKeywords.push(keyword);
    }
  }
  return foundKeywords;
}
/**
 * Extracts text content from HTML body
 * @param htmlContent - HTML content to extract text from
 * @returns Plain text content
 */
function extractTextFromHtml(htmlContent) {
  if (!htmlContent) return '';
  // Remove HTML tags and decode common HTML entities
  const textContent = htmlContent.replace(/<[^>]*>/g, ' ') // Remove HTML tags
  .replace(/&nbsp;/g, ' ') // Replace non-breaking spaces
  .replace(/&amp;/g, '&') // Decode ampersands
  .replace(/&lt;/g, '<') // Decode less than
  .replace(/&gt;/g, '>') // Decode greater than
  .replace(/&quot;/g, '"') // Decode quotes
  .replace(/&#39;/g, "'") // Decode apostrophes
  .replace(/\s+/g, ' ') // Replace multiple whitespace with single space
  .trim();
  return textContent;
}
/**
 * Handler for message send events (emails)
 * @param event - The Office.js event object
 */
function onMessageSendHandler(event) {
  console.log('Message send handler triggered');
  // Get the current item (email being sent)
  const item = Office.context.mailbox.item;
  if (!item) {
    console.error('No item found');
    event.completed({
      allowEvent: true
    });
    return;
  }
  // Get the body of the email
  item.body.getAsync(Office.CoercionType.Html, result => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get email body:', result.error);
      // Allow sending if we can't get the body
      event.completed({
        allowEvent: true
      });
      return;
    }
    const htmlContent = result.value;
    const textContent = extractTextFromHtml(htmlContent);
    console.log('Scanning email content for sensitive keywords...');
    const foundKeywords = scanForSensitiveKeywords(textContent);
    if (foundKeywords.length > 0) {
      console.log('Sensitive keywords found:', foundKeywords);
      const keywordList = foundKeywords.join(', ');
      const errorMessage = `This email contains sensitive keywords (${keywordList}). Please review before sending.`;
      // Block the send and show notification
      event.completed({
        allowEvent: false,
        errorMessage: errorMessage
      });
    } else {
      console.log('No sensitive keywords found, allowing send');
      // Allow the email to be sent
      event.completed({
        allowEvent: true
      });
    }
  });
}
/**
 * Handler for appointment send events (calendar invitations)
 * @param event - The Office.js event object
 */
function onAppointmentSendHandler(event) {
  console.log('Appointment send handler triggered');
  // Get the current item (appointment being sent)
  const item = Office.context.mailbox.item;
  if (!item) {
    console.error('No item found');
    event.completed({
      allowEvent: true
    });
    return;
  }
  // Get the body of the appointment
  item.body.getAsync(Office.CoercionType.Html, result => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get appointment body:', result.error);
      // Allow sending if we can't get the body
      event.completed({
        allowEvent: true
      });
      return;
    }
    const htmlContent = result.value;
    const textContent = extractTextFromHtml(htmlContent);
    // Also check the subject line for appointments
    const subject = item.subject || '';
    const fullContent = `${subject} ${textContent}`;
    console.log('Scanning appointment content for sensitive keywords...');
    const foundKeywords = scanForSensitiveKeywords(fullContent);
    if (foundKeywords.length > 0) {
      console.log('Sensitive keywords found:', foundKeywords);
      const keywordList = foundKeywords.join(', ');
      const errorMessage = `This appointment contains sensitive keywords (${keywordList}). Please review before sending.`;
      // Block the send and show notification
      event.completed({
        allowEvent: false,
        errorMessage: errorMessage
      });
    } else {
      console.log('No sensitive keywords found, allowing send');
      // Allow the appointment to be sent
      event.completed({
        allowEvent: true
      });
    }
  });
}
// Register the event handlers
(function () {
  "use strict";

  Office.onReady(() => {
    console.log('Office.js is ready, registering event handlers');
    // Make handlers available globally for manifest
    window.onMessageSendHandler = onMessageSendHandler;
    window.onAppointmentSendHandler = onAppointmentSendHandler;
  });
})();
// Export for potential use in other modules

/******/ })()
;
//# sourceMappingURL=commands.js.map