/**
 * Storage Module
 * Handles session storage for security and fresh capture validation
 */

import { CONFIG } from './config.js';

/**
 * Generate random session ID
 */
export function generateSessionId() {
    return `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Generate random nonce
 */
export function generateNonce() {
    return `nonce_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Initialize session
 * Creates a new session with unique ID
 */
export function initSession() {
    const sessionId = generateSessionId();
    sessionStorage.setItem(CONFIG.SESSION.STORAGE_KEY, sessionId);
    return sessionId;
}

/**
 * Get current session ID
 */
export function getSessionId() {
    let sessionId = sessionStorage.getItem(CONFIG.SESSION.STORAGE_KEY);
    if (!sessionId) {
        sessionId = initSession();
    }
    return sessionId;
}

/**
 * Clear session
 */
export function clearSession() {
    sessionStorage.removeItem(CONFIG.SESSION.STORAGE_KEY);
    sessionStorage.removeItem(CONFIG.SESSION.CAPTURE_KEY);
    sessionStorage.removeItem(CONFIG.SESSION.TIMESTAMP_KEY);
    sessionStorage.removeItem(CONFIG.SESSION.NONCE_KEY);
}

/**
 * Store captured photo data
 * Only stores in sessionStorage (clears on page refresh)
 */
export function storeCapturedPhoto(photoData) {
    const timestamp = Date.now();
    const nonce = generateNonce();
    
    sessionStorage.setItem(CONFIG.SESSION.CAPTURE_KEY, photoData);
    sessionStorage.setItem(CONFIG.SESSION.TIMESTAMP_KEY, timestamp.toString());
    sessionStorage.setItem(CONFIG.SESSION.NONCE_KEY, nonce);
    
    return { timestamp, nonce };
}

/**
 * Get captured photo data
 */
export function getCapturedPhoto() {
    const photoData = sessionStorage.getItem(CONFIG.SESSION.CAPTURE_KEY);
    const timestamp = sessionStorage.getItem(CONFIG.SESSION.TIMESTAMP_KEY);
    const nonce = sessionStorage.getItem(CONFIG.SESSION.NONCE_KEY);
    
    if (!photoData || !timestamp || !nonce) {
        return null;
    }
    
    return {
        photoData,
        timestamp: parseInt(timestamp, 10),
        nonce
    };
}

/**
 * Clear captured photo data
 * Forces fresh capture
 */
export function clearCapturedPhoto() {
    sessionStorage.removeItem(CONFIG.SESSION.CAPTURE_KEY);
    sessionStorage.removeItem(CONFIG.SESSION.TIMESTAMP_KEY);
    sessionStorage.removeItem(CONFIG.SESSION.NONCE_KEY);
}

/**
 * Validate photo freshness
 * Checks if photo was captured in this session
 */
export function validatePhotoFreshness() {
    const captureData = getCapturedPhoto();
    
    if (!captureData) {
        return false;
    }
    
    // Check if timestamp is recent (within last 5 minutes)
    const now = Date.now();
    const age = now - captureData.timestamp;
    const maxAge = 5 * 60 * 1000; // 5 minutes
    
    return age < maxAge;
}

/**
 * Check if session is valid
 */
export function isSessionValid() {
    const sessionId = sessionStorage.getItem(CONFIG.SESSION.STORAGE_KEY);
    return !!sessionId;
}

/**
 * Store attendance state temporarily
 */
export function storeAttendanceState(employeeName, hasClockIn, hasClockOut) {
    const state = {
        employeeName,
        hasClockIn,
        hasClockOut,
        timestamp: Date.now()
    };
    sessionStorage.setItem('attendance_state', JSON.stringify(state));
}

/**
 * Get attendance state
 */
export function getAttendanceState() {
    const stateStr = sessionStorage.getItem('attendance_state');
    if (!stateStr) return null;
    
    try {
        const state = JSON.parse(stateStr);
        
        // Check if state is recent (within last 30 seconds)
        const now = Date.now();
        const age = now - state.timestamp;
        if (age > 30000) {
            // State is too old, clear it
            clearAttendanceState();
            return null;
        }
        
        return state;
    } catch (error) {
        console.error('Error parsing attendance state:', error);
        return null;
    }
}

/**
 * Clear attendance state
 */
export function clearAttendanceState() {
    sessionStorage.removeItem('attendance_state');
}

/**
 * Setup storage event listener for cross-tab communication
 */
export function setupStorageListener(callback) {
    window.addEventListener('storage', (event) => {
        if (event.key === CONFIG.SESSION.STORAGE_KEY) {
            callback(event);
        }
    });
}

/**
 * Check if localStorage is available
 */
export function isStorageAvailable() {
    try {
        const test = '__storage_test__';
        sessionStorage.setItem(test, test);
        sessionStorage.removeItem(test);
        return true;
    } catch (error) {
        return false;
    }
}

/**
 * Get all session data (for debugging)
 */
export function getSessionData() {
    return {
        sessionId: getSessionId(),
        capturedPhoto: !!getCapturedPhoto(),
        photoFresh: validatePhotoFreshness(),
        attendanceState: getAttendanceState()
    };
}
