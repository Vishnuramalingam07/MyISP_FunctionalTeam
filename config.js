/**
 * MyISP Tools - API Configuration
 * 
 * This file configures the backend API URL for all tools.
 * 
 * When running locally:
 *   - Set API_BASE_URL to '' (empty string) or 'http://localhost:5000'
 * 
 * When deployed:
 *   - Set API_BASE_URL to your deployed backend URL
 *   - Example: 'https://your-server.com' or 'http://server-ip:5000'
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Backend API base URL
 * 
 * OPTIONS:
 * 1. Local development: '' or 'http://localhost:5000'
 * 2. Azure deployment: 'https://myisp-tools.azurewebsites.net'
 * 3. Custom server: 'http://your-server-ip:5000'
 */
const API_BASE_URL = ''; // Use relative URLs (works for localhost and dev tunnels)

// ============================================================================
// HELPER FUNCTION
// ============================================================================

/**
 * Constructs full API endpoint URL
 * @param {string} endpoint - API endpoint path (e.g., '/api/generate-report')
 * @returns {string} Full URL to the API endpoint
 */
function getApiUrl(endpoint) {
    // Remove leading slash if present to avoid double slashes
    const cleanEndpoint = endpoint.startsWith('/') ? endpoint : '/' + endpoint;
    
    // If API_BASE_URL is empty, return relative path (for local development)
    if (!API_BASE_URL || API_BASE_URL === '') {
        return cleanEndpoint;
    }
    
    // Remove trailing slash from base URL to avoid double slashes
    const cleanBaseUrl = API_BASE_URL.endsWith('/') 
        ? API_BASE_URL.slice(0, -1) 
        : API_BASE_URL;
    
    return cleanBaseUrl + cleanEndpoint;
}
