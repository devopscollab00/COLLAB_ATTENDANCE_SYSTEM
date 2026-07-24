/**
 * Modal Module
 * Handles all modal dialogs (loading, success, error, confirm)
 */

import { CONFIG } from './config.js';

/**
 * Show loading modal
 */
export function showLoading(message = CONFIG.MESSAGES.LOADING.DEFAULT) {
    const modal = document.getElementById('loadingModal');
    const loadingText = document.getElementById('loadingText');
    
    if (modal && loadingText) {
        loadingText.textContent = message;
        modal.classList.add('active');
        
        // Disable body scroll
        document.body.style.overflow = 'hidden';
    }
}

/**
 * Hide loading modal
 */
export function hideLoading() {
    const modal = document.getElementById('loadingModal');
    
    if (modal) {
        modal.classList.remove('active');
        
        // Enable body scroll
        document.body.style.overflow = '';
    }
}

/**
 * Show success modal
 */
export function showSuccess(title, message, onClose = null) {
    const modal = document.getElementById('successModal');
    const titleElement = document.getElementById('successTitle');
    const messageElement = document.getElementById('successMessage');
    const okBtn = document.getElementById('successOkBtn');
    
    if (modal && titleElement && messageElement && okBtn) {
        titleElement.textContent = title;
        messageElement.textContent = message;
        modal.classList.add('active');
        
        // Disable body scroll
        document.body.style.overflow = 'hidden';
        
        // Setup close handler
        const closeHandler = () => {
            hideSuccess();
            if (onClose) onClose();
            okBtn.removeEventListener('click', closeHandler);
        };
        
        okBtn.addEventListener('click', closeHandler);
        
        // Auto-hide after duration
        if (CONFIG.UI.AUTO_HIDE_SUCCESS > 0) {
            setTimeout(() => {
                if (modal.classList.contains('active')) {
                    closeHandler();
                }
            }, CONFIG.UI.AUTO_HIDE_SUCCESS);
        }
    }
}

/**
 * Hide success modal
 */
export function hideSuccess() {
    const modal = document.getElementById('successModal');
    
    if (modal) {
        modal.classList.remove('active');
        
        // Enable body scroll
        document.body.style.overflow = '';
    }
}

/**
 * Show error modal
 */
export function showError(title, message, onClose = null) {
    const modal = document.getElementById('errorModal');
    const titleElement = document.getElementById('errorTitle');
    const messageElement = document.getElementById('errorMessage');
    const okBtn = document.getElementById('errorOkBtn');
    
    if (modal && titleElement && messageElement && okBtn) {
        titleElement.textContent = title;
        messageElement.textContent = message;
        modal.classList.add('active');
        
        // Disable body scroll
        document.body.style.overflow = 'hidden';
        
        // Setup close handler
        const closeHandler = () => {
            hideError();
            if (onClose) onClose();
            okBtn.removeEventListener('click', closeHandler);
        };
        
        okBtn.addEventListener('click', closeHandler);
    }
}

/**
 * Hide error modal
 */
export function hideError() {
    const modal = document.getElementById('errorModal');
    
    if (modal) {
        modal.classList.remove('active');
        
        // Enable body scroll
        document.body.style.overflow = '';
    }
}

/**
 * Show confirm modal
 */
export function showConfirm(title, message, onConfirm, onCancel = null) {
    const modal = document.getElementById('confirmModal');
    const titleElement = document.getElementById('confirmTitle');
    const messageElement = document.getElementById('confirmMessage');
    const okBtn = document.getElementById('confirmOkBtn');
    const cancelBtn = document.getElementById('confirmCancelBtn');
    
    if (modal && titleElement && messageElement && okBtn && cancelBtn) {
        titleElement.textContent = title;
        messageElement.textContent = message;
        modal.classList.add('active');
        
        // Disable body scroll
        document.body.style.overflow = 'hidden';
        
        // Setup confirm handler
        const confirmHandler = () => {
            hideConfirm();
            if (onConfirm) onConfirm();
            okBtn.removeEventListener('click', confirmHandler);
            cancelBtn.removeEventListener('click', cancelHandler);
        };
        
        // Setup cancel handler
        const cancelHandler = () => {
            hideConfirm();
            if (onCancel) onCancel();
            okBtn.removeEventListener('click', confirmHandler);
            cancelBtn.removeEventListener('click', cancelHandler);
        };
        
        okBtn.addEventListener('click', confirmHandler);
        cancelBtn.addEventListener('click', cancelHandler);
    }
}

/**
 * Hide confirm modal
 */
export function hideConfirm() {
    const modal = document.getElementById('confirmModal');
    
    if (modal) {
        modal.classList.remove('active');
        
        // Enable body scroll
        document.body.style.overflow = '';
    }
}

/**
 * Close modal on backdrop click
 */
export function setupModalBackdropClose() {
    const modals = [
        'loadingModal',
        'successModal',
        'errorModal',
        'confirmModal'
    ];
    
    modals.forEach(modalId => {
        const modal = document.getElementById(modalId);
        if (modal) {
            modal.addEventListener('click', (e) => {
                // Only close if clicking the backdrop, not the content
                if (e.target === modal) {
                    // Don't close loading modal on backdrop click
                    if (modalId !== 'loadingModal') {
                        modal.classList.remove('active');
                        document.body.style.overflow = '';
                    }
                }
            });
        }
    });
}

/**
 * Close modal on Escape key
 */
export function setupModalEscapeClose() {
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            // Find active modal
            const activeModal = document.querySelector('.modal.active');
            
            if (activeModal) {
                // Don't close loading modal with Escape
                if (activeModal.id !== 'loadingModal') {
                    activeModal.classList.remove('active');
                    document.body.style.overflow = '';
                }
            }
        }
    });
}

/**
 * Initialize modal system
 */
export function initModals() {
    setupModalBackdropClose();
    setupModalEscapeClose();
}

/**
 * Check if any modal is open
 */
export function isModalOpen() {
    return document.querySelector('.modal.active') !== null;
}

/**
 * Close all modals
 */
export function closeAllModals() {
    hideLoading();
    hideSuccess();
    hideError();
    hideConfirm();
}
