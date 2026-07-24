/**
 * Configuration Module
 * Central configuration for the application
 */

export const CONFIG = {
    // API Configuration
    API: {
        ENDPOINT: 'https://script.google.com/macros/s/AKfycbzzcWQCdHFSN9Ik8KCe4kDGrusvyAyc6t6jYLff2bNeUxdT6kZ3lSBiYsMeqv3xzfpj/exec',
        TIMEOUT: 30000, // 30 seconds
        RETRY_ATTEMPTS: 3,
        RETRY_DELAY: 2000 // 2 seconds
    },

    // Camera Configuration
    CAMERA: {
        WIDTH: 1280,
        HEIGHT: 720,
        FACING_MODE: 'user', // or 'environment' for back camera
        QUALITY: 0.9 // JPEG quality
    },

    // Session Configuration
    SESSION: {
        STORAGE_KEY: 'attendance_session',
        CAPTURE_KEY: 'attendance_capture',
        TIMESTAMP_KEY: 'attendance_timestamp',
        NONCE_KEY: 'attendance_nonce'
    },

    // Location Dropdown Options
    LOCATIONS: [
        'Main Entrance',
        'Warehouse',
        'Office',
        'Production',
        'Lobby',
        'Admin Building',
        'Conference Room',
        'Parking Area'
    ],

    // Validation Rules
    VALIDATION: {
        EMPLOYEE_REQUIRED: true,
        LOCATION_REQUIRED: true,
        PHOTO_REQUIRED: true,
        MIN_PHOTO_SIZE: 1024, // 1KB minimum
        MAX_PHOTO_SIZE: 5242880 // 5MB maximum
    },

    // UI Configuration
    UI: {
        TOAST_DURATION: 5000, // 5 seconds
        LOADING_MIN_DISPLAY: 500, // Minimum loading display time
        DEBOUNCE_DELAY: 300, // Debounce delay for inputs
        AUTO_HIDE_SUCCESS: 3000 // Auto hide success modal
    },

    // Messages
    MESSAGES: {
        LOADING: {
            DEFAULT: 'Processing...',
            TIME_IN: 'Recording Time In...',
            TIME_OUT: 'Recording Time Out...',
            CHECKING_STATUS: 'Checking attendance status...',
            LOADING_DATA: 'Loading data...'
        },
        SUCCESS: {
            TIME_IN: 'Time In recorded successfully!',
            TIME_OUT: 'Time Out recorded successfully!'
        },
        ERROR: {
            NETWORK: 'Network error. Please check your connection.',
            TIMEOUT: 'Request timeout. Please try again.',
            API_ERROR: 'Server error. Please try again later.',
            VALIDATION: 'Please fill in all required fields.',
            CAMERA_DENIED: 'Camera access denied. Please enable camera permissions.',
            CAMERA_NOT_FOUND: 'No camera found on this device.',
            PHOTO_REQUIRED: 'Please capture a photo before submitting.',
            ALREADY_CLOCKED_IN: 'You have already clocked in today.',
            ALREADY_CLOCKED_OUT: 'You have already clocked out today.',
            NOT_CLOCKED_IN: 'You need to clock in first.',
            INVALID_SESSION: 'Invalid session. Please refresh the page.',
            UNKNOWN: 'An unexpected error occurred. Please try again.'
        },
        CONFIRM: {
            TIME_IN: 'Confirm Time In?',
            TIME_OUT: 'Confirm Time Out?'
        }
    },

    // Greeting Messages
    GREETINGS: {
        MORNING: 'Good morning! Ready to start your day?',
        AFTERNOON: 'Good afternoon! Hope your day is going well.',
        EVENING: 'Good evening! Almost time to wrap up.',
        NIGHT: 'Working late? Stay safe!'
    }
};

/**
 * Get greeting based on current time
 */
export function getGreeting() {
    const hour = new Date().getHours();
    
    if (hour >= 5 && hour < 12) {
        return CONFIG.GREETINGS.MORNING;
    } else if (hour >= 12 && hour < 17) {
        return CONFIG.GREETINGS.AFTERNOON;
    } else if (hour >= 17 && hour < 21) {
        return CONFIG.GREETINGS.EVENING;
    } else {
        return CONFIG.GREETINGS.NIGHT;
    }
}

/**
 * Get status text based on attendance state
 */
export function getStatusText(hasClockIn, hasClockOut) {
    if (hasClockOut) {
        return 'Completed - You have clocked out';
    } else if (hasClockIn) {
        return 'Clocked In - Ready to clock out';
    } else {
        return 'Ready - Please clock in';
    }
}
