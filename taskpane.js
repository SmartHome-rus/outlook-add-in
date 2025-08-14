/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./node_modules/html-loader/dist/runtime/getUrl.js":
/*!*********************************************************!*\
  !*** ./node_modules/html-loader/dist/runtime/getUrl.js ***!
  \*********************************************************/
/***/ ((module) => {



module.exports = function (url, options) {
  if (!options) {
    // eslint-disable-next-line no-param-reassign
    options = {};
  }

  if (!url) {
    return url;
  } // eslint-disable-next-line no-underscore-dangle, no-param-reassign


  url = String(url.__esModule ? url.default : url);

  if (options.hash) {
    // eslint-disable-next-line no-param-reassign
    url += options.hash;
  }

  if (options.maybeNeedQuotes && /[\t\n\f\r "'=<>`]/.test(url)) {
    return "\"".concat(url, "\"");
  }

  return url;
};

/***/ }),

/***/ "./src/commands/commands.ts":
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

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


/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ ((module, __unused_webpack_exports, __webpack_require__) => {

module.exports = __webpack_require__.p + "74fd0c8ee4bf7ec6ebd8.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
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
/******/ 	/* webpack/runtime/global */
/******/ 	(() => {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
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
/******/ 	/* webpack/runtime/publicPath */
/******/ 	(() => {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	(() => {
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
(() => {
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _commands_commands__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../commands/commands */ "./src/commands/commands.ts");
/**
 * Taskpane functionality for the Sensitive Data Scanner
 */
/// <reference path="../types/office.d.ts" />
// Import the scanning functionality from commands

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
 * Test the scanner on the current email
 */
function testCurrentEmail() {
  const resultDiv = document.getElementById('test-result');
  if (!resultDiv) return;
  // Clear previous results
  resultDiv.innerHTML = '';
  resultDiv.className = '';
  const item = Office.context.mailbox.item;
  if (!item) {
    resultDiv.innerHTML = '<div class="test-result-error">No email is currently open for testing.</div>';
    return;
  }
  // Show loading state
  resultDiv.innerHTML = '<div class="test-result-info">Scanning current email...</div>';
  // Get the body of the current item
  item.body.getAsync(Office.CoercionType.Html, result => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      resultDiv.innerHTML = '<div class="test-result-error">Failed to read email content.</div>';
      return;
    }
    const htmlContent = result.value;
    const textContent = extractTextFromHtml(htmlContent);
    // Also get subject if available
    const subject = item.subject || '';
    const fullContent = `${subject} ${textContent}`;
    const foundKeywords = (0,_commands_commands__WEBPACK_IMPORTED_MODULE_0__.scanForSensitiveKeywords)(fullContent);
    if (foundKeywords.length > 0) {
      const keywordList = foundKeywords.join(', ');
      resultDiv.innerHTML = `<div class="test-result-warning">⚠️ Sensitive keywords found: <strong>${keywordList}</strong><br>This email would be blocked from sending.</div>`;
    } else {
      resultDiv.innerHTML = '<div class="test-result-success">✅ No sensitive keywords detected. This email would be allowed to send.</div>';
    }
  });
}
/**
 * Initialize the taskpane
 */
function initialize() {
  // Check if Office is ready
  if (typeof Office !== 'undefined') {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Bind the test scanner button
    const testButton = document.getElementById('test-scanner');
    if (testButton) {
      testButton.addEventListener('click', testCurrentEmail);
    }
  } else {
    // Show sideload message if Office is not available
    document.getElementById("sideload-msg").style.display = "flex";
    document.getElementById("app-body").style.display = "none";
  }
}
// Initialize when Office is ready
Office.onReady(info => {
  initialize();
});
// Also initialize immediately in case Office is already ready
if (typeof Office !== 'undefined') {
  initialize();
}
})();

// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
(() => {
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../node_modules/html-loader/dist/runtime/getUrl.js */ "./node_modules/html-loader/dist/runtime/getUrl.js");
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__);
// Imports

var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
// Module
var ___HTML_LOADER_REPLACEMENT_0___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_0___);
var code = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Sensitive Data Scanner</title>\r\n    \r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js\"><" + "/script>\r\n    \r\n    <!-- Local CSS -->\r\n    <link rel=\"stylesheet\" href=\"" + ___HTML_LOADER_REPLACEMENT_0___ + "\" />\r\n</head>\r\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\r\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\r\n        <div class=\"logo-placeholder\">🛡️</div>\r\n        <h1 class=\"ms-font-su\">Sensitive Data Scanner</h1>\r\n    </header>\r\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\r\n        <h2 class=\"ms-font-xl\">Please sideload your add-in to see app body.</h2>\r\n    </section>\r\n    <main id=\"app-body\" class=\"ms-welcome__main\" style=\"display: none;\">\r\n        <h2 class=\"ms-font-xl\">Protect your sensitive information</h2>\r\n        <div class=\"feature-section\">\r\n            <p class=\"ms-font-m\">This add-in automatically scans your emails and calendar invitations for sensitive keywords before sending.</p>\r\n            \r\n            <h3 class=\"ms-font-l\">Protected Keywords:</h3>\r\n            <ul id=\"keywords-list\" class=\"keywords-list\">\r\n                <li>confidential</li>\r\n                <li>internal use only</li>\r\n                <li>secret</li>\r\n                <li>project-alpha</li>\r\n                <li>classified</li>\r\n                <li>restricted</li>\r\n                <li>proprietary</li>\r\n                <li>do not distribute</li>\r\n                <li>sensitive</li>\r\n                <li>private and confidential</li>\r\n            </ul>\r\n            \r\n            <div class=\"status-section\">\r\n                <h3 class=\"ms-font-l\">Scanner Status:</h3>\r\n                <p id=\"scanner-status\" class=\"ms-font-m status-active\">✅ Active - Monitoring outgoing messages</p>\r\n            </div>\r\n            \r\n            <div class=\"test-section\">\r\n                <h3 class=\"ms-font-l\">Test Scanner:</h3>\r\n                <p class=\"ms-font-s\">Try composing an email with the word \"confidential\" to test the scanner.</p>\r\n                <button id=\"test-scanner\" class=\"ms-Button ms-Button--primary\">\r\n                    <span class=\"ms-Button-label\">Test Current Email</span>\r\n                </button>\r\n                <div id=\"test-result\" style=\"margin-top: 10px;\"></div>\r\n            </div>\r\n        </div>\r\n    </main>\r\n</body>\r\n</html>\r\n";
// Exports
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (code);
})();

/******/ })()
;
//# sourceMappingURL=taskpane.js.map