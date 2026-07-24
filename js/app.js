/**
 * Main Application Module
 * Initializes and coordinates all modules
 */

import { CONFIG } from './config.js';
import * as API from './api.js';
import * as Storage from './storage.js';
import * as Camera from './camera.js';
import * as Validation from './validation.js';
import * as Modal from './modal.js';
import * as UI from './ui.js';

// Application state
let currentEmployee = null;
let currentLocation = null;
let hasClockIn = false;
let hasClockOut = false;
let isSubmitting = false;

/**
 * Initialize application
 */
async function init() {
    console.log('Initializing Employee Attendance System...');
    
    // Initialize session
    Storage.initSession();
    
    // Initialize UI
    UI.initUI();
    Modal.initModals();
    
    // Setup event listeners
    setupEventListeners();
    
    // Load initial data
    await loadInitialData();
    
    // Setup connection monitor
    setupConnectionMonitor();
    
    // Check camera availability
    checkCameraAvailability();
    
    console.log('Application initialized successfully');
}

/**
 * Load initial data (employees, locations)
 */
async function loadInitialData() {
    try {
        Modal.showLoading(CONFIG.MESSAGES.LOADING.LOADING_DATA);
        
        // Load employees and locations in parallel
        const [employees, locations] = await Promise.all([
            API.loadEmployees(),
            API.loadLocations()
        ]);
        
        // Populate dropdowns
        const employeeSelect = document.getElementById('employeeName');
        const locationSelect = document.getElementById('location');
        
        if (employeeSelect) {
            UI.populateSelect(employeeSelect, employees, 'Select your name');
        }
        
        if (locationSelect) {
            UI.populateSelect(locationSelect, locations, 'Select location');
        }
        
        Modal.hideLoading();
        
    } catch (error) {
        console.error('Error loading initial data:', error);
        Modal.hideLoading();
        UI.showToast('error', 'Error', 'Failed to load data. Please refresh the page.');
    }
}

/**
 * Setup all event listeners
 */
function setupEventListeners() {
    // Employee selection change
    const employeeSelect = document.getElementById('employeeName');
    if (employeeSelect) {
        employeeSelect.addEventListener('change', handleEmployeeChange);
    }
    
    // Location selection change
    const locationSelect = document.getElementById('location');
    if (locationSelect) {
        locationSelect.addEventListener('change', handleLocationChange);
    }
    
    // Camera controls
    const openCameraBtn = document.getElementById('openCameraBtn');
    if (openCameraBtn) {
        openCameraBtn.addEventListener('click', handleOpenCamera);
    }
    
    const captureBtn = document.getElementById('captureBtn');
    if (captureBtn) {
        captureBtn.addEventListener('click', handleCapturePhoto);
    }
    
    const retakeBtn = document.getElementById('retakeBtn');
    if (retakeBtn) {
        retakeBtn.addEventListener('click', handleRetakePhoto);
    }
    
    // Action buttons
    const timeInBtn = document.getElementById('timeInBtn');
    if (timeInBtn) {
        timeInBtn.addEventListener('click', handleTimeIn);
    }
    
    const timeOutBtn = document.getElementById('timeOutBtn');
    if (timeOutBtn) {
        timeOutBtn.addEventListener('click', handleTimeOut);
    }
    
    // Prevent double submission
    const form = document.getElementById('attendanceForm');
    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
        });
    }
}

/**
 * Handle employee selection change
 */
async function handleEmployeeChange(event) {
    currentEmployee = event.target.value;
    Validation.clearValidationErrors();
    
    if (!currentEmployee) {
        UI.hideStatus();
        hasClockIn = false;
        hasClockOut = false;
        updateUI();
        return;
    }
    
    // Check attendance status
    try {
        Modal.showLoading(CONFIG.MESSAGES.LOADING.CHECKING_STATUS);
        
        const status = await API.checkAttendanceStatus(currentEmployee);
        hasClockIn = status.hasClockIn;
        hasClockOut = status.hasClockOut;
        
        // Show status
        UI.showStatus(hasClockIn, hasClockOut);
        
        // Reset camera if needed for Time Out
        if (hasClockIn && !hasClockOut) {
            Camera.resetCamera();
            UI.updateCameraControls('idle');
            UI.showToast('info', 'Time Out Required', 'Please capture a new photo for Time Out');
        }
        
        // If already clocked out, show message
        if (hasClockOut) {
            UI.showToast('success', 'Completed', 'You have already completed attendance for today');
        }
        
        Modal.hideLoading();
        updateUI();
        
    } catch (error) {
        console.error('Error checking attendance:', error);
        Modal.hideLoading();
        UI.showToast('error', 'Error', 'Failed to check attendance status');
    }
}

/**
 * Handle location selection change
 */
function handleLocationChange(event) {
    currentLocation = event.target.value;
    Validation.clearValidationErrors();
    updateUI();
}

/**
 * Handle open camera
 */
async function handleOpenCamera() {
    try {
        UI.updateCameraControls('opening');
        await Camera.openCamera();
        UI.updateCameraControls('camera_open');
        UI.showToast('success', 'Camera Ready', 'Camera is ready. Click capture when ready.');
    } catch (error) {
        console.error('Error opening camera:', error);
        UI.showToast('error', 'Camera Error', error.message);
        UI.updateCameraControls('idle');
    }
}

/**
 * Handle capture photo
 */
async function handleCapturePhoto() {
    try {
        await Camera.capturePhoto();
        UI.updateCameraControls('captured');
        UI.showToast('success', 'Photo Captured', 'Photo captured successfully!');
        updateUI();
    } catch (error) {
        console.error('Error capturing photo:', error);
        UI.showToast('error', 'Capture Error', 'Failed to capture photo. Please try again.');
    }
}

/**
 * Handle retake photo
 */
function handleRetakePhoto() {
    Camera.resetCamera();
    UI.updateCameraControls('idle');
    updateUI();
}

/**
 * Handle Time In
 */
async function handleTimeIn() {
    // Prevent double submission
    if (isSubmitting) {
        console.log('Already submitting...');
        return;
    }
    
    // Clear previous errors
    Validation.clearValidationErrors();
    
    // Validate form
    const validation = Validation.validateForm(currentEmployee, currentLocation);
    if (!validation.isValid) {
        Validation.displayValidationErrors(validation.errors);
        UI.showToast('error', 'Validation Error', CONFIG.MESSAGES.ERROR.VALIDATION);
        UI.shakeElement(document.getElementById('attendanceForm'));
        return;
    }
    
    // Get captured photo
    const capturedPhoto = Storage.getCapturedPhoto();
    if (!capturedPhoto) {
        UI.showToast('error', 'Photo Required', CONFIG.MESSAGES.ERROR.PHOTO_REQUIRED);
        return;
    }
    
    // Confirm action
    Modal.showConfirm(
        CONFIG.MESSAGES.CONFIRM.TIME_IN,
        `Record Time In for ${currentEmployee} at ${currentLocation}?`,
        async () => {
            await submitTimeIn(capturedPhoto.photoData);
        }
    );
}

/**
 * Submit Time In
 */
async function submitTimeIn(photoData) {
    isSubmitting = true;
    
    try {
        Modal.showLoading(CONFIG.MESSAGES.LOADING.TIME_IN);
        UI.disableButtons();
        
        // Submit to API
        const response = await API.submitTimeIn(
            currentEmployee,
            currentLocation,
            photoData
        );
        
        Modal.hideLoading();
        
        // Show success
        Modal.showSuccess(
            'Time In Successful',
            CONFIG.MESSAGES.SUCCESS.TIME_IN,
            () => {
                // Reset form after success
                resetForm();
                hasClockIn = true;
                UI.showStatus(hasClockIn, hasClockOut);
                updateUI();
            }
        );
        
        // Clear captured photo
        Storage.clearCapturedPhoto();
        Camera.resetCamera();
        
    } catch (error) {
        console.error('Error submitting time in:', error);
        Modal.hideLoading();
        Modal.showError('Time In Failed', error.message || CONFIG.MESSAGES.ERROR.API_ERROR);
        
    } finally {
        isSubmitting = false;
        UI.enableButtons();
        updateUI();
    }
}

/**
 * Handle Time Out
 */
async function handleTimeOut() {
    // Prevent double submission
    if (isSubmitting) {
        console.log('Already submitting...');
        return;
    }
    
    // Clear previous errors
    Validation.clearValidationErrors();
    
    // Validate form
    const validation = Validation.validateForm(currentEmployee, currentLocation);
    if (!validation.isValid) {
        Validation.displayValidationErrors(validation.errors);
        UI.showToast('error', 'Validation Error', CONFIG.MESSAGES.ERROR.VALIDATION);
        UI.shakeElement(document.getElementById('attendanceForm'));
        return;
    }
    
    // Get captured photo
    const capturedPhoto = Storage.getCapturedPhoto();
    if (!capturedPhoto) {
        UI.showToast('error', 'Photo Required', 'Please capture a fresh photo for Time Out');
        return;
    }
    
    // Confirm action
    Modal.showConfirm(
        CONFIG.MESSAGES.CONFIRM.TIME_OUT,
        `Record Time Out for ${currentEmployee} at ${currentLocation}?`,
        async () => {
            await submitTimeOut(capturedPhoto.photoData);
        }
    );
}

/**
 * Submit Time Out
 */
async function submitTimeOut(photoData) {
    isSubmitting = true;
    
    try {
        Modal.showLoading(CONFIG.MESSAGES.LOADING.TIME_OUT);
        UI.disableButtons();
        
        // Submit to API
        const response = await API.submitTimeOut(
            currentEmployee,
            currentLocation,
            photoData
        );
        
        Modal.hideLoading();
        
        // Show success
        Modal.showSuccess(
            'Time Out Successful',
            CONFIG.MESSAGES.SUCCESS.TIME_OUT,
            () => {
                // Reset form after success
                resetForm();
                hasClockOut = true;
                UI.showStatus(hasClockIn, hasClockOut);
                updateUI();
            }
        );
        
        // Clear captured photo
        Storage.clearCapturedPhoto();
        Camera.resetCamera();
        
    } catch (error) {
        console.error('Error submitting time out:', error);
        Modal.hideLoading();
        Modal.showError('Time Out Failed', error.message || CONFIG.MESSAGES.ERROR.API_ERROR);
        
    } finally {
        isSubmitting = false;
        UI.enableButtons();
        updateUI();
    }
}

/**
 * Update UI based on current state
 */
function updateUI() {
    const hasCapturedPhoto = Camera.isReadyToSubmit();
    UI.updateButtonStates(hasClockIn, hasClockOut, hasCapturedPhoto);
}

/**
 * Reset form
 */
function resetForm() {
    // Clear selections
    const employeeSelect = document.getElementById('employeeName');
    const locationSelect = document.getElementById('location');
    
    if (employeeSelect) employeeSelect.value = '';
    if (locationSelect) locationSelect.value = '';
    
    currentEmployee = null;
    currentLocation = null;
    
    // Reset camera
    Camera.resetCamera();
    UI.updateCameraControls('idle');
    
    // Clear validation errors
    Validation.clearValidationErrors();
    
    // Clear status
    UI.hideStatus();
    
    // Reset state
    hasClockIn = false;
    hasClockOut = false;
}

/**
 * Setup connection monitor
 */
function setupConnectionMonitor() {
    API.setupConnectionMonitor(
        () => {
            UI.showToast('success', 'Online', 'Connection restored');
        },
        () => {
            UI.showToast('error', 'Offline', CONFIG.MESSAGES.ERROR.NETWORK);
        }
    );
}

/**
 * Check camera availability
 */
function checkCameraAvailability() {
    if (!Camera.isCameraAvailable()) {
        UI.showToast('warning', 'Camera Unavailable', 'Camera is not available on this device');
        
        // Disable camera buttons
        const openCameraBtn = document.getElementById('openCameraBtn');
        if (openCameraBtn) {
            openCameraBtn.disabled = true;
            openCameraBtn.textContent = 'Camera Not Available';
        }
    }
}

/**
 * Handle page visibility change
 */
document.addEventListener('visibilitychange', () => {
    if (!document.hidden) {
        // Page became visible
        console.log('Page visible - checking session');
        
        // Validate session
        if (!Storage.isSessionValid()) {
            console.log('Session invalid - reinitializing');
            Storage.initSession();
        }
    }
});

/**
 * Handle page unload
 */
window.addEventListener('beforeunload', (event) => {
    // Stop camera if running
    Camera.stopCamera();
    
    // Warn if submission in progress
    if (isSubmitting) {
        event.preventDefault();
        event.returnValue = 'Submission in progress. Are you sure you want to leave?';
        return event.returnValue;
    }
});

/**
 * Handle errors globally
 */
window.addEventListener('error', (event) => {
    console.error('Global error:', event.error);
    
    // Don't show error modal for script load errors
    if (event.error && event.error.message) {
        UI.showToast('error', 'Error', 'An unexpected error occurred');
    }
});

/**
 * Handle unhandled promise rejections
 */
window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled promise rejection:', event.reason);
    UI.showToast('error', 'Error', 'An unexpected error occurred');
});

// Initialize application when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}

// Export for debugging
window.__app = {
    API,
    Storage,
    Camera,
    Validation,
    Modal,
    UI,
    CONFIG,
    getState: () => ({
        currentEmployee,
        currentLocation,
        hasClockIn,
        hasClockOut,
        isSubmitting,
        cameraState: Camera.getCameraState(),
        sessionData: Storage.getSessionData()
    })
};
