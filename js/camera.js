/**
 * Camera Module
 * Handles live camera access and photo capture with security
 */

import { CONFIG } from './config.js';
import { storeCapturedPhoto, clearCapturedPhoto } from './storage.js';
import { showToast } from './ui.js';

// Camera state machine
const CameraState = {
    IDLE: 'idle',
    CAMERA_OPEN: 'camera_open',
    CAPTURED: 'captured',
    READY_TO_SUBMIT: 'ready_to_submit'
};

let currentState = CameraState.IDLE;
let mediaStream = null;

/**
 * Get current camera state
 */
export function getCameraState() {
    return currentState;
}

/**
 * Set camera state
 */
function setCameraState(newState) {
    console.log(`Camera state transition: ${currentState} -> ${newState}`);
    currentState = newState;
}

/**
 * Check if camera is available
 */
export function isCameraAvailable() {
    return !!(navigator.mediaDevices && navigator.mediaDevices.getUserMedia);
}

/**
 * Request camera permission and open camera
 */
export async function openCamera() {
    if (!isCameraAvailable()) {
        throw new Error(CONFIG.MESSAGES.ERROR.CAMERA_NOT_FOUND);
    }

    // Check if camera is already open
    if (currentState === CameraState.CAMERA_OPEN) {
        console.log('Camera already open');
        return;
    }

    try {
        // Stop any existing stream first
        if (mediaStream) {
            stopCamera();
        }

        // Request camera access
        const constraints = {
            video: {
                width: { ideal: CONFIG.CAMERA.WIDTH },
                height: { ideal: CONFIG.CAMERA.HEIGHT },
                facingMode: CONFIG.CAMERA.FACING_MODE
            },
            audio: false
        };

        mediaStream = await navigator.mediaDevices.getUserMedia(constraints);
        
        // Get video element
        const videoElement = document.getElementById('cameraPreview');
        if (!videoElement) {
            throw new Error('Video element not found');
        }

        // Set video source
        videoElement.srcObject = mediaStream;
        
        // Wait for video to be ready
        await new Promise((resolve, reject) => {
            videoElement.onloadedmetadata = () => {
                videoElement.play()
                    .then(resolve)
                    .catch(reject);
            };
            
            // Timeout after 5 seconds
            setTimeout(() => reject(new Error('Video load timeout')), 5000);
        });

        // Update state
        setCameraState(CameraState.CAMERA_OPEN);
        
        return true;

    } catch (error) {
        console.error('Error opening camera:', error);
        
        // Handle specific error types
        if (error.name === 'NotAllowedError' || error.name === 'PermissionDeniedError') {
            throw new Error(CONFIG.MESSAGES.ERROR.CAMERA_DENIED);
        } else if (error.name === 'NotFoundError' || error.name === 'DevicesNotFoundError') {
            throw new Error(CONFIG.MESSAGES.ERROR.CAMERA_NOT_FOUND);
        } else {
            throw new Error(`Camera error: ${error.message}`);
        }
    }
}

/**
 * Capture photo from video stream
 */
export async function capturePhoto() {
    if (currentState !== CameraState.CAMERA_OPEN) {
        throw new Error('Camera is not open');
    }

    try {
        const videoElement = document.getElementById('cameraPreview');
        const canvas = document.getElementById('canvas');
        const photoPreview = document.getElementById('photoPreview');

        if (!videoElement || !canvas || !photoPreview) {
            throw new Error('Required elements not found');
        }

        // Set canvas size to match video
        canvas.width = videoElement.videoWidth;
        canvas.height = videoElement.videoHeight;

        // Draw current video frame to canvas
        const ctx = canvas.getContext('2d');
        ctx.drawImage(videoElement, 0, 0, canvas.width, canvas.height);

        // Add timestamp overlay
        addTimestampOverlay(ctx, canvas.width, canvas.height);

        // Convert to base64
        const photoData = canvas.toDataURL('image/jpeg', CONFIG.CAMERA.QUALITY);

        // Validate photo size
        const photoSize = photoData.length;
        if (photoSize < CONFIG.VALIDATION.MIN_PHOTO_SIZE) {
            throw new Error('Photo is too small');
        }
        if (photoSize > CONFIG.VALIDATION.MAX_PHOTO_SIZE) {
            throw new Error('Photo is too large');
        }

        // Store captured photo with timestamp and nonce
        storeCapturedPhoto(photoData);

        // Display preview
        photoPreview.src = photoData;
        photoPreview.style.display = 'block';

        // Stop camera stream immediately
        stopCamera();

        // Update state
        setCameraState(CameraState.CAPTURED);

        return photoData;

    } catch (error) {
        console.error('Error capturing photo:', error);
        throw error;
    }
}

/**
 * Add timestamp overlay to photo
 */
function addTimestampOverlay(ctx, width, height) {
    const now = new Date();
    const timestamp = now.toLocaleString('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
    });

    // Set overlay style
    ctx.fillStyle = 'rgba(0, 0, 0, 0.6)';
    ctx.fillRect(0, height - 40, width, 40);

    // Set text style
    ctx.fillStyle = '#FFFFFF';
    ctx.font = 'bold 14px Inter, sans-serif';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';

    // Draw timestamp
    ctx.fillText(`📅 ${timestamp}`, 10, height - 20);
}

/**
 * Stop camera stream
 */
export function stopCamera() {
    if (mediaStream) {
        // Stop all tracks
        mediaStream.getTracks().forEach(track => {
            track.stop();
        });
        
        // Clear video element
        const videoElement = document.getElementById('cameraPreview');
        if (videoElement) {
            videoElement.srcObject = null;
        }
        
        mediaStream = null;
        console.log('Camera stopped');
    }

    // Update state if camera was open
    if (currentState === CameraState.CAMERA_OPEN) {
        setCameraState(CameraState.IDLE);
    }
}

/**
 * Reset camera (clear captured photo and reset state)
 */
export function resetCamera() {
    // Stop camera if running
    stopCamera();
    
    // Clear captured photo
    clearCapturedPhoto();
    
    // Hide photo preview
    const photoPreview = document.getElementById('photoPreview');
    if (photoPreview) {
        photoPreview.style.display = 'none';
        photoPreview.src = '';
    }
    
    // Reset state
    setCameraState(CameraState.IDLE);
}

/**
 * Mark photo as ready to submit
 */
export function markReadyToSubmit() {
    if (currentState === CameraState.CAPTURED) {
        setCameraState(CameraState.READY_TO_SUBMIT);
    }
}

/**
 * Check if photo is ready to submit
 */
export function isReadyToSubmit() {
    return currentState === CameraState.READY_TO_SUBMIT || currentState === CameraState.CAPTURED;
}

/**
 * Cleanup on page unload
 */
window.addEventListener('beforeunload', () => {
    stopCamera();
});

/**
 * Cleanup on visibility change (tab switch)
 */
document.addEventListener('visibilitychange', () => {
    if (document.hidden && currentState === CameraState.CAMERA_OPEN) {
        // Optionally stop camera when tab is hidden
        // stopCamera();
    }
});

/**
 * Check camera permissions status
 */
export async function checkCameraPermission() {
    if (!navigator.permissions) {
        // Permissions API not supported
        return 'unsupported';
    }

    try {
        const result = await navigator.permissions.query({ name: 'camera' });
        return result.state; // 'granted', 'denied', or 'prompt'
    } catch (error) {
        console.error('Error checking camera permission:', error);
        return 'unsupported';
    }
}

/**
 * Get available cameras (front/back)
 */
export async function getAvailableCameras() {
    if (!navigator.mediaDevices || !navigator.mediaDevices.enumerateDevices) {
        return [];
    }

    try {
        const devices = await navigator.mediaDevices.enumerateDevices();
        return devices.filter(device => device.kind === 'videoinput');
    } catch (error) {
        console.error('Error getting cameras:', error);
        return [];
    }
}
