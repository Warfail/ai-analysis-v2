// frontend/assets/script.js
// Shared functions and configuration

const API_BASE = 'http://localhost:8000/api';

// Utility Functions
function showAlert(message, type = 'success', duration = 5000) {
    // Create or find alert container
    let alertContainer = document.getElementById('alertContainer');
    if (!alertContainer) {
        alertContainer = document.createElement('div');
        alertContainer.id = 'alertContainer';
        alertContainer.style.position = 'fixed';
        alertContainer.style.top = '20px';
        alertContainer.style.right = '20px';
        alertContainer.style.zIndex = '1000';
        document.body.appendChild(alertContainer);
    }
    
    // Create alert
    const alert = document.createElement('div');
    alert.className = `alert alert-${type}`;
    alert.textContent = message;
    alert.style.marginBottom = '10px';
    alert.style.display = 'block';
    
    alertContainer.appendChild(alert);
    
    // Remove after duration
    setTimeout(() => {
        alert.remove();
        if (alertContainer.children.length === 0) {
            alertContainer.remove();
        }
    }, duration);
}

function formatDate(dateString) {
    if (!dateString) return 'Unknown date';
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
}

function showLoading(show, elementId = 'loadingIndicator') {
    const loading = document.getElementById(elementId);
    if (loading) {
        loading.style.display = show ? 'block' : 'none';
    }
}

async function checkBackendHealth() {
    try {
        const response = await axios.get(`${API_BASE}/`);
        const status = document.getElementById('aiStatus');
        if (status) {
            if (response.data.status === 'running') {
                status.textContent = 'AI Active';
                status.classList.add('active');
            }
        }
        return true;
    } catch {
        const status = document.getElementById('aiStatus');
        if (status) {
            status.textContent = 'Backend Unavailable';
            status.classList.remove('active');
        }
        return false;
    }
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    checkBackendHealth();
    
    // Set today's date in date inputs
    const today = new Date().toISOString().split('T')[0];
    document.querySelectorAll('input[type="date"]').forEach(input => {
        if (!input.value) {
            input.value = today;
        }
    });
    
    // Highlight active nav link
    const currentPage = window.location.pathname.split('/').pop() || 'index.html';
    document.querySelectorAll('.nav-link').forEach(link => {
        const href = link.getAttribute('href');
        if (href === currentPage || (currentPage === '' && href === 'index.html')) {
            link.classList.add('active');
        }
    });
});