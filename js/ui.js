/**
 * UI Module
 * Handles UI updates, animations, and toast notifications
 */

import { CONFIG, getGreeting, getStatusText } from './config.js';

/**
 * Initialize UI
 */
export function initUI() {
    updateDateTime();
    updateGreeting();
    
    // Update time every second
    setInterval(updateDateTime, 1000);
}

/**
 * Update date and time displays
 */
function updateDateTime() {
    const now = new Date();
    
    // Update live time
    const timeDisplay = document.getElementById('timeDisplay');
    if (timeDisplay) {
        const hours = now.getHours();
        const minutes = now.getMinutes();
        const seconds = now.getSeconds();
        const ampm = hours >= 12 ? 'PM' : 'AM';
        const displayHours = hours % 12 || 12;
        
        timeDisplay.textContent = 
            `${pad(displayHours)}:${pad(minutes)}:${pad(seconds)} ${ampm}`;
    }
    
    // Update date display
    const dateDisplay = document.getElementById('currentDate');
    if (dateDisplay) {
        const options = { 
            weekday: 'long', 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric' 
        };
        dateDisplay.textContent = now.toLocaleDateString('en-US', options);
    }
}

/**
 * Pad numbers with leading zero
 */
function pad(num) {
    return num.toString().padStart(2, '0');
}

/**
 * Update greeting based on time
 */
function updateGreeting() {
    const greetingText = document.getElementById('greetingText');
    if (greetingText) {
        greetingText.textContent = getGreeting();
    }
}

/**
 * Show status badge
 */
export function showStatus(hasClockIn, hasBreakStart, hasBreakEnd, hasClockOut) {
    const statusBadge = document.getElementById('statusBadge');
    const statusText = document.getElementById('statusText');
    
    if (statusBadge && statusText) {
        statusText.textContent = getStatusText(hasClockIn, hasBreakStart, hasBreakEnd, hasClockOut);
        statusBadge.style.display = 'flex';
        
        // Add animation
        statusBadge.classList.add('fade-in-up');
        
        // Update badge color based on status
        if (hasClockOut) {
            statusBadge.style.background = 'linear-gradient(135deg, rgba(213, 0, 0, 0.1), rgba(213, 0, 0, 0.05))';
            statusBadge.style.borderColor = 'rgba(213, 0, 0, 0.2)';
            statusBadge.style.color = 'var(--danger-dark)';
        } else if (hasBreakStart && !hasBreakEnd) {
            statusBadge.style.background = 'linear-gradient(135deg, rgba(255, 111, 0, 0.1), rgba(255, 111, 0, 0.05))';
            statusBadge.style.borderColor = 'rgba(255, 111, 0, 0.2)';
            statusBadge.style.color = 'var(--warning-dark)';
        } else if (hasClockIn) {
            statusBadge.style.background = 'linear-gradient(135deg, rgba(0, 200, 83, 0.1), rgba(0, 200, 83, 0.05))';
            statusBadge.style.borderColor = 'rgba(0, 200, 83, 0.2)';
            statusBadge.style.color = 'var(--success-dark)';
        }
    }
}

/**
 * Hide status badge
 */
export function hideStatus() {
    const statusBadge = document.getElementById('statusBadge');
    if (statusBadge) {
        statusBadge.style.display = 'none';
    }
}

/**
 * Update button states based on attendance status
 */
export function updateButtonStates(hasClockIn, hasBreakStart, hasBreakEnd, hasClockOut, hasPhoto) {
    const timeInBtn = document.getElementById('timeInBtn');
    const breakStartBtn = document.getElementById('breakStartBtn');
    const breakEndBtn = document.getElementById('breakEndBtn');
    const timeOutBtn = document.getElementById('timeOutBtn');
    
    if (timeInBtn && breakStartBtn && breakEndBtn && timeOutBtn) {
        if (hasClockOut) {
            // Already clocked out - disable all
            timeInBtn.disabled = true;
            timeInBtn.style.display = 'none';
            breakStartBtn.disabled = true;
            breakStartBtn.style.display = 'none';
            breakEndBtn.disabled = true;
            breakEndBtn.style.display = 'none';
            timeOutBtn.disabled = true;
            timeOutBtn.style.display = 'flex';
        } else if (hasClockIn) {
            // Clocked in - show break and time out buttons
            timeInBtn.disabled = true;
            timeInBtn.style.display = 'none';
            
            if (!hasBreakStart) {
                // Show break start button (no photo needed)
                breakStartBtn.disabled = false;
                breakStartBtn.style.display = 'flex';
                breakEndBtn.style.display = 'none';
            } else if (hasBreakStart && !hasBreakEnd) {
                // Show break end button (no photo needed)
                breakStartBtn.style.display = 'none';
                breakEndBtn.disabled = false;
                breakEndBtn.style.display = 'flex';
            } else {
                // Break completed, hide break buttons
                breakStartBtn.style.display = 'none';
                breakEndBtn.style.display = 'none';
            }
            
            // Show time out button (needs photo)
            timeOutBtn.disabled = !hasPhoto;
            timeOutBtn.style.display = 'flex';
        } else {
            // Not clocked in - show only time in button
            timeInBtn.disabled = !hasPhoto;
            timeInBtn.style.display = 'flex';
            breakStartBtn.style.display = 'none';
            breakEndBtn.style.display = 'none';
            timeOutBtn.disabled = true;
            timeOutBtn.style.display = 'none';
        }
    }
}

/**
 * Show camera controls based on state
 */
export function updateCameraControls(state) {
    const openCameraBtn = document.getElementById('openCameraBtn');
    const captureBtn = document.getElementById('captureBtn');
    const retakeBtn = document.getElementById('retakeBtn');
    const cameraOverlay = document.getElementById('cameraOverlay');
    const cameraPreview = document.getElementById('cameraPreview');
    const photoPreview = document.getElementById('photoPreview');
    
    switch (state) {
        case 'idle':
            if (openCameraBtn) openCameraBtn.style.display = 'flex';
            if (captureBtn) captureBtn.style.display = 'none';
            if (retakeBtn) retakeBtn.style.display = 'none';
            if (cameraOverlay) cameraOverlay.style.display = 'flex';
            if (cameraPreview) cameraPreview.style.display = 'none';
            if (photoPreview) photoPreview.style.display = 'none';
            break;
            
        case 'camera_open':
            if (openCameraBtn) openCameraBtn.style.display = 'none';
            if (captureBtn) captureBtn.style.display = 'flex';
            if (retakeBtn) retakeBtn.style.display = 'none';
            if (cameraOverlay) cameraOverlay.style.display = 'none';
            if (cameraPreview) cameraPreview.style.display = 'block';
            if (photoPreview) photoPreview.style.display = 'none';
            break;
            
        case 'captured':
            if (openCameraBtn) openCameraBtn.style.display = 'none';
            if (captureBtn) captureBtn.style.display = 'none';
            if (retakeBtn) retakeBtn.style.display = 'flex';
            if (cameraOverlay) cameraOverlay.style.display = 'none';
            if (cameraPreview) cameraPreview.style.display = 'none';
            if (photoPreview) photoPreview.style.display = 'block';
            break;
    }
}

/**
 * Toast notification system
 */
let toastCounter = 0;

export function showToast(type, title, message, duration = CONFIG.UI.TOAST_DURATION) {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    
    const toastId = `toast-${toastCounter++}`;
    
    // Create toast element
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.id = toastId;
    toast.innerHTML = `
        <i class="toast-icon bi ${getToastIcon(type)}"></i>
        <div class="toast-content">
            <div class="toast-title">${title}</div>
            <div class="toast-message">${message}</div>
        </div>
        <button class="toast-close" onclick="document.getElementById('${toastId}').remove()">
            <i class="bi bi-x"></i>
        </button>
    `;
    
    // Add to container
    container.appendChild(toast);
    
    // Add animation
    toast.classList.add('slide-in-right');
    
    // Auto remove after duration
    setTimeout(() => {
        if (toast && toast.parentElement) {
            toast.classList.add('removing');
            setTimeout(() => {
                if (toast && toast.parentElement) {
                    toast.remove();
                }
            }, 300);
        }
    }, duration);
}

/**
 * Get icon for toast type
 */
function getToastIcon(type) {
    switch (type) {
        case 'success':
            return 'bi-check-circle-fill';
        case 'error':
            return 'bi-x-circle-fill';
        case 'warning':
            return 'bi-exclamation-triangle-fill';
        case 'info':
        default:
            return 'bi-info-circle-fill';
    }
}

/**
 * Enable all buttons
 */
export function enableButtons() {
    const buttons = document.querySelectorAll('.btn');
    buttons.forEach(btn => {
        if (!btn.hasAttribute('data-keep-disabled')) {
            btn.disabled = false;
        }
    });
}

/**
 * Disable all buttons
 */
export function disableButtons() {
    const buttons = document.querySelectorAll('.btn');
    buttons.forEach(btn => {
        btn.disabled = true;
    });
}

/**
 * Add ripple effect to button
 */
export function addRippleEffect() {
    const buttons = document.querySelectorAll('.btn');
    
    buttons.forEach(button => {
        button.addEventListener('click', function(e) {
            const ripple = document.createElement('span');
            const rect = this.getBoundingClientRect();
            const size = Math.max(rect.width, rect.height);
            const x = e.clientX - rect.left - size / 2;
            const y = e.clientY - rect.top - size / 2;
            
            ripple.style.width = ripple.style.height = size + 'px';
            ripple.style.left = x + 'px';
            ripple.style.top = y + 'px';
            ripple.classList.add('ripple');
            
            this.appendChild(ripple);
            
            setTimeout(() => {
                ripple.remove();
            }, 600);
        });
    });
}

/**
 * Animate element
 */
export function animateElement(element, animationClass) {
    if (!element) return;
    
    element.classList.add(animationClass);
    
    element.addEventListener('animationend', () => {
        element.classList.remove(animationClass);
    }, { once: true });
}

/**
 * Shake element (for errors)
 */
export function shakeElement(element) {
    if (!element) return;
    animateElement(element, 'shake');
}

/**
 * Skeleton loader
 */
export function showSkeleton(element) {
    if (!element) return;
    element.classList.add('skeleton');
}

export function hideSkeleton(element) {
    if (!element) return;
    element.classList.remove('skeleton');
}

/**
 * Update card title
 */
export function updateCardTitle(title) {
    const cardTitle = document.querySelector('.card-title');
    if (cardTitle) {
        cardTitle.textContent = title;
        animateElement(cardTitle, 'fade-in');
    }
}

/**
 * Update card subtitle
 */
export function updateCardSubtitle(subtitle) {
    const cardSubtitle = document.querySelector('.card-subtitle');
    if (cardSubtitle) {
        cardSubtitle.textContent = subtitle;
        animateElement(cardSubtitle, 'fade-in');
    }
}

/**
 * Scroll to element
 */
export function scrollToElement(element, behavior = 'smooth') {
    if (!element) return;
    
    element.scrollIntoView({
        behavior: behavior,
        block: 'center'
    });
}

/**
 * Focus on element
 */
export function focusElement(element) {
    if (!element) return;
    
    setTimeout(() => {
        element.focus();
    }, 100);
}

/**
 * Populate select dropdown
 */
export function populateSelect(selectElement, options, placeholder = 'Select option') {
    if (!selectElement) return;
    
    // Clear existing options
    selectElement.innerHTML = '';
    
    // Add placeholder
    const placeholderOption = document.createElement('option');
    placeholderOption.value = '';
    placeholderOption.textContent = placeholder;
    selectElement.appendChild(placeholderOption);
    
    // Add options
    options.forEach(option => {
        const optionElement = document.createElement('option');
        optionElement.value = option;
        optionElement.textContent = option;
        selectElement.appendChild(optionElement);
    });
}

/**
 * Show/hide element
 */
export function showElement(element) {
    if (!element) return;
    element.style.display = '';
}

export function hideElement(element) {
    if (!element) return;
    element.style.display = 'none';
}
