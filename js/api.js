/**
 * API Module
 * Handles all communication with Google Apps Script backend
 */

import { CONFIG } from './config.js';
import { showToast } from './ui.js';

/**
 * Check if online
 */
export function isOnline() {
    return navigator.onLine;
}

/**
 * Generic API request handler with retry logic
 */
async function apiRequest(payload, retryCount = 0) {
    // Check if online
    if (!isOnline()) {
        throw new Error(CONFIG.MESSAGES.ERROR.NETWORK);
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), CONFIG.API.TIMEOUT);

    try {
        const response = await fetch(CONFIG.API.ENDPOINT, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: new URLSearchParams(payload).toString(),
            signal: controller.signal,
            mode: 'cors'
        });

        clearTimeout(timeoutId);

        // Check if response is OK
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        // Parse response
        const text = await response.text();
        
        // Try to parse as JSON first, fallback to text
        try {
            return JSON.parse(text);
        } catch {
            // If not JSON, return as text response
            return { success: true, message: text };
        }

    } catch (error) {
        clearTimeout(timeoutId);

        // Handle abort (timeout)
        if (error.name === 'AbortError') {
            if (retryCount < CONFIG.API.RETRY_ATTEMPTS) {
                await sleep(CONFIG.API.RETRY_DELAY);
                return apiRequest(payload, retryCount + 1);
            }
            throw new Error(CONFIG.MESSAGES.ERROR.TIMEOUT);
        }

        // Handle network errors with retry
        if (retryCount < CONFIG.API.RETRY_ATTEMPTS) {
            await sleep(CONFIG.API.RETRY_DELAY);
            return apiRequest(payload, retryCount + 1);
        }

        // Rethrow error
        throw error;
    }
}

/**
 * Sleep utility for retry delays
 */
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Validate Employee ID and get employee details
 * Calls backend getEmployeeByID function
 */
export async function validateEmployeeId(employeeId) {
    try {
        // Call Google Apps Script function directly
        const url = `${CONFIG.API.ENDPOINT}?action=getEmployeeByID&employeeId=${encodeURIComponent(employeeId)}`;
        
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), CONFIG.API.TIMEOUT);

        const response = await fetch(url, {
            method: 'GET',
            signal: controller.signal,
            mode: 'cors'
        });

        clearTimeout(timeoutId);

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        
        if (result.success === false) {
            throw new Error(result.message || 'Invalid Employee ID');
        }

        return {
            success: true,
            employeeId: result.employeeId,
            employeeName: result.employeeName,
            fullName: result.fullName
        };

    } catch (error) {
        console.error('Error validating employee ID:', error);
        
        if (error.name === 'AbortError') {
            throw new Error('Request timeout. Please try again.');
        }
        
        throw error;
    }
}

/**
 * Load employees from backend
 * @deprecated - Not needed anymore, using manual ID entry
 */
export async function loadEmployees() {
    try {
        const response = await fetch(CONFIG.API.ENDPOINT, {
            method: 'GET',
            mode: 'cors'
        });

        if (!response.ok) {
            throw new Error('Failed to load employees');
        }

        const html = await response.text();
        
        // Parse employees from HTML
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const datalist = doc.querySelector('#employeeList');
        
        if (!datalist) {
            throw new Error('Employee list not found');
        }

        const options = datalist.querySelectorAll('option');
        const employees = Array.from(options).map(opt => opt.value).filter(v => v);
        
        return employees;

    } catch (error) {
        console.error('Error loading employees:', error);
        showToast('warning', 'Warning', 'Could not load employee list. Please refresh the page.');
        return [];
    }
}

/**
 * Load locations from backend
 */
export async function loadLocations() {
    try {
        const response = await fetch(CONFIG.API.ENDPOINT, {
            method: 'GET',
            mode: 'cors'
        });

        if (!response.ok) {
            throw new Error('Failed to load locations');
        }

        const html = await response.text();
        
        // Parse locations from HTML
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const datalist = doc.querySelector('#locationList');
        
        if (!datalist) {
            // Return default locations if not found in HTML
            return CONFIG.LOCATIONS;
        }

        const options = datalist.querySelectorAll('option');
        const locations = Array.from(options).map(opt => opt.value).filter(v => v);
        
        return locations.length > 0 ? locations : CONFIG.LOCATIONS;

    } catch (error) {
        console.error('Error loading locations:', error);
        // Return default locations on error
        return CONFIG.LOCATIONS;
    }
}

/**
 * Check attendance status for an employee
 */
export async function checkAttendanceStatus(employeeName) {
    try {
        const payload = {
            action: 'checkStatus',
            employeeName: employeeName
        };

        const response = await apiRequest(payload);
        
        return {
            hasClockIn: response.hasClockIn || false,
            hasClockOut: response.hasClockOut || false
        };

    } catch (error) {
        console.error('Error checking attendance status:', error);
        throw new Error(CONFIG.MESSAGES.ERROR.API_ERROR);
    }
}

/**
 * Submit Time In
 */
export async function submitTimeIn(employeeName, location, photoBase64) {
    try {
        const payload = {
            type: 'IN',
            fullName: employeeName,
            location: location,
            photoBase64: photoBase64
        };

        const response = await apiRequest(payload);
        
        if (response.success === false) {
            throw new Error(response.message || CONFIG.MESSAGES.ERROR.API_ERROR);
        }

        return response;

    } catch (error) {
        console.error('Error submitting time in:', error);
        throw error;
    }
}

/**
 * Submit Time Out
 */
export async function submitTimeOut(employeeName, location, photoBase64) {
    try {
        const payload = {
            type: 'OUT',
            fullName: employeeName,
            location: location,
            photoBase64: photoBase64
        };

        const response = await apiRequest(payload);
        
        if (response.success === false) {
            throw new Error(response.message || CONFIG.MESSAGES.ERROR.API_ERROR);
        }

        return response;

    } catch (error) {
        console.error('Error submitting time out:', error);
        throw error;
    }
}

/**
 * Ping API to check if it's available
 */
export async function pingAPI() {
    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 5000); // 5 second timeout for ping

        const response = await fetch(CONFIG.API.ENDPOINT, {
            method: 'GET',
            signal: controller.signal,
            mode: 'cors'
        });

        clearTimeout(timeoutId);
        return response.ok;

    } catch (error) {
        console.error('API ping failed:', error);
        return false;
    }
}

/**
 * Setup online/offline event listeners
 */
export function setupConnectionMonitor(onOnline, onOffline) {
    window.addEventListener('online', () => {
        console.log('Connection restored');
        if (onOnline) onOnline();
    });

    window.addEventListener('offline', () => {
        console.log('Connection lost');
        if (onOffline) onOffline();
    });
}
