/**
 * Sensitive Data Scanner - On-Send Event Handlers
 * This file contains the core logic for scanning emails and appointments
 * for sensitive keywords and blocking sending when found.
 */
/**
 * Scans text content for sensitive keywords
 * @param content - The text content to scan
 * @returns Array of found keywords
 */
declare function scanForSensitiveKeywords(content: string): string[];
/**
 * Handler for message send events (emails)
 * @param event - The Office.js event object
 */
declare function onMessageSendHandler(event: any): void;
/**
 * Handler for appointment send events (calendar invitations)
 * @param event - The Office.js event object
 */
declare function onAppointmentSendHandler(event: any): void;
export { onMessageSendHandler, onAppointmentSendHandler, scanForSensitiveKeywords };
