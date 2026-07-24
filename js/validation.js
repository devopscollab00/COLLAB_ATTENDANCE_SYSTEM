/**
 * Validation Module
 * Handles form validation and input checking
 */

import { CONFIG } from './config.js';
import { getCapturedPhoto, validatePhotoFreshness } from './storage.js';
import { isReadyToSubmit } from './camera.js';

/**
 * Validation result object
 */
class ValidationResult {
    constructor(isValid, errors = {}) {
        this.isValid = isValid;
        this.errors = errors;
    }
}

/**
 * Validate employee ID input
 */
export function validateEmployeeId(employeeId) {
    if (!employeeId || employeeId.trim() === '') {
        return new ValidationResult(false, {
            employee: 'Please enter your Employee ID'
        });
    }
    
    // Basic format validation (adjust as needed)
    if (employeeId.trim().length < 2) {
        return new ValidationResult(false, {
            employee: 'Employee ID must be at least 2 characters'
        });
    }
    
    return new ValidationResult(true);
}

/**
 * Validate employee selection
 * @deprecated - Use validateEmployeeId instead
 */
export function validateEmployee(employeeName) {
    if (!employeeName || employeeName.trim() === '') {
        return new ValidationResult(false, {
            employee: 'Please select your name'
        });
    }
    return new ValidationResult(true);
}

/**
 * Validate location selection
 */
export function validateLocation(location) {
    if (!location || location.trim() === '') {
        return new ValidationResult(false, {
            location: 'Please select a location'
        });
    }
    return new ValidationResult(true);
}

/**
 * Validate photo capture
 */
export function validatePhoto() {
    const capturedPhoto = getCapturedPhoto();
    
    if (!capturedPhoto) {
        return new ValidationResult(false, {
            photo: CONFIG.MESSAGES.ERROR.PHOTO_REQUIRED
        });
    }

    // Check if photo is fresh
    if (!validatePhotoFreshness()) {
        return new ValidationResult(false, {
            photo: 'Photo is too old. Please capture a new photo.'
        });
    }

    // Check if ready to submit
    if (!isReadyToSubmit()) {
        return new ValidationResult(false, {
            photo: 'Photo is not ready. Please capture a photo first.'
        });
    }

    return new ValidationResult(true);
}

/**
 * Validate entire form
 */
export function validateForm(employeeName, location) {
    const errors = {};
    let isValid = true;

    // Validate employee (check if validated - currentEmployee should be set)
    if (!employeeName) {
        isValid = false;
        errors.employee = 'Please enter and validate your Employee ID';
    }

    // Validate location
    const locationValidation = validateLocation(location);
    if (!locationValidation.isValid) {
        isValid = false;
        Object.assign(errors, locationValidation.errors);
    }

    // Validate photo
    const photoValidation = validatePhoto();
    if (!photoValidation.isValid) {
        isValid = false;
        Object.assign(errors, photoValidation.errors);
    }

    return new ValidationResult(isValid, errors);
}

/**
 * Display validation errors in the UI
 */
export function displayValidationErrors(errors) {
    // Clear all error displays first
    clearValidationErrors();

    // Display employee error
    if (errors.employee) {
        const employeeInput = document.getElementById('employeeName');
        const employeeError = document.getElementById('employeeError');
        if (employeeInput && employeeError) {
            employeeInput.classList.add('error');
            employeeError.textContent = errors.employee;
            employeeError.classList.add('active');
        }
    }

    // Display location error
    if (errors.location) {
        const locationInput = document.getElementById('location');
        const locationError = document.getElementById('locationError');
        if (locationInput && locationError) {
            locationInput.classList.add('error');
            locationError.textContent = errors.location;
            locationError.classList.add('active');
        }
    }

    // Display photo error
    if (errors.photo) {
        const photoError = document.getElementById('photoError');
        if (photoError) {
            photoError.textContent = errors.photo;
            photoError.classList.add('active');
        }
    }
}

/**
 * Clear all validation error displays
 */
export function clearValidationErrors() {
    // Clear employee error
    const employeeInput = document.getElementById('employeeName');
    const employeeError = document.getElementById('employeeError');
    if (employeeInput) employeeInput.classList.remove('error');
    if (employeeError) {
        employeeError.textContent = '';
        employeeError.classList.remove('active');
    }

    // Clear location error
    const locationInput = document.getElementById('location');
    const locationError = document.getElementById('locationError');
    if (locationInput) locationInput.classList.remove('error');
    if (locationError) {
        locationError.textContent = '';
        locationError.classList.remove('active');
    }

    // Clear photo error
    const photoError = document.getElementById('photoError');
    if (photoError) {
        photoError.textContent = '';
        photoError.classList.remove('active');
    }
}

/**
 * Validate email format (utility)
 */
export function validateEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

/**
 * Validate phone format (utility)
 */
export function validatePhone(phone) {
    const phoneRegex = /^[\d\s\-\+\(\)]+$/;
    return phoneRegex.test(phone) && phone.replace(/\D/g, '').length >= 10;
}

/**
 * Sanitize input (prevent XSS)
 */
export function sanitizeInput(input) {
    if (typeof input !== 'string') return input;
    
    const div = document.createElement('div');
    div.textContent = input;
    return div.innerHTML;
}

/**
 * Validate string length
 */
export function validateLength(str, min, max) {
    const length = str ? str.trim().length : 0;
    return length >= min && length <= max;
}

/**
 * Check if input contains only allowed characters
 */
export function validateCharacters(str, allowedPattern) {
    return allowedPattern.test(str);
}

/**
 * Real-time validation helper
 */
export function setupRealtimeValidation(inputElement, validationFn, errorElement) {
    if (!inputElement || !errorElement) return;

    inputElement.addEventListener('blur', () => {
        const value = inputElement.value;
        const result = validationFn(value);
        
        if (!result.isValid) {
            inputElement.classList.add('error');
            errorElement.textContent = Object.values(result.errors)[0] || 'Invalid input';
            errorElement.classList.add('active');
        } else {
            inputElement.classList.remove('error');
            errorElement.textContent = '';
            errorElement.classList.remove('active');
        }
    });

    inputElement.addEventListener('input', () => {
        // Clear error on input
        inputElement.classList.remove('error');
        errorElement.textContent = '';
        errorElement.classList.remove('active');
    });
}

/**
 * Debounce function for real-time validation
 */
export function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

/**
 * Check for empty or whitespace-only strings
 */
export function isEmpty(str) {
    return !str || str.trim().length === 0;
}

/**
 * Validate form data before submission
 */
export function validateBeforeSubmit(formData) {
    const errors = [];

    if (isEmpty(formData.employeeName)) {
        errors.push('Employee name is required');
    }

    if (isEmpty(formData.location)) {
        errors.push('Location is required');
    }

    if (!formData.photoBase64) {
        errors.push('Photo is required');
    }

    return {
        isValid: errors.length === 0,
        errors: errors
    };
}
