/* global Office */

const N8N_BASE_URL = "https://workflows.prostarpics.com/webhook-test";

class AuthService {
  constructor() {
    this.token = null;
    this.tokenExpiry = null;
  }

  isAuthenticated() {
    if (!this.token || !this.tokenExpiry) {
      return false;
    }
    // Check if token expires in next 5 minutes
    return Date.now() < (this.tokenExpiry - 300000);
  }

  async getToken() {
    if (this.isAuthenticated()) {
      return this.token;
    }

    // Get fresh token from Microsoft
    return new Promise((resolve, reject) => {
      Office.context.auth.getAccessTokenAsync(
        {
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: false // We're using it for our own backend, not MS Graph
        },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            this.token = result.value;
            
            // Decode JWT to get expiry (token is base64 encoded in 3 parts)
            try {
              const payload = JSON.parse(atob(result.value.split('.')[1]));
              this.tokenExpiry = payload.exp * 1000; // Convert to milliseconds
            } catch (e) {
              // Fallback: assume 1 hour expiry
              this.tokenExpiry = Date.now() + (60 * 60 * 1000);
            }
            
            resolve(this.token);
          } else {
            // Handle different error codes (based on Microsoft's documentation)
            let errorMessage = 'Authentication failed';
            
            if (result.error.code === 13001) {
              errorMessage = 'No one is signed into Office. Please sign in and try again.';
            } else if (result.error.code === 13002) {
              errorMessage = 'User cancelled the consent prompt.';
            } else if (result.error.code === 13003) {
              errorMessage = 'User type is not supported. Contact your administrator.';
            } else if (result.error.code === 13004) {
              errorMessage = 'Resource not available. Try again later.';
            } else if (result.error.code === 13005) {
              errorMessage = 'SSO is not supported for this account type.';
            } else if (result.error.code === 13006) {
              errorMessage = 'Office on the web is experiencing issues. Close browser and restart.';
            } else if (result.error.code === 13007) {
              errorMessage = 'Add-in is not registered correctly. Contact your administrator.';
            } else if (result.error.code === 13008) {
              errorMessage = 'Office is still processing. Wait and try again.';
            } else if (result.error.code === 13009) {
              errorMessage = 'Platform does not support this version of Office.js.';
            } else if (result.error.code === 13010) {
              errorMessage = 'Browser zone configuration issue. Contact your administrator.';
            } else if (result.error.code === 13012) {
              errorMessage = 'Add-in running in unsupported context.';
            } else {
              errorMessage = `Authentication error ${result.error.code}: ${result.error.message}`;
            }
            
            reject(new Error(errorMessage));
          }
        }
      );
    });
  }

  logout() {
    // With SSO, we can't really "logout" - user is authenticated to Office
    // But we can clear our cached token to force refresh
    this.token = null;
    this.tokenExpiry = null;
  }

  async makeAuthenticatedRequest(endpoint, options = {}) {
    try {
      const token = await this.getToken();
      
      const response = await fetch(`${N8N_BASE_URL}${endpoint}`, {
        ...options,
        headers: {
          ...options.headers,
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        }
      });

      if (response.status === 401) {
        // Token invalid, clear and retry once
        this.logout();
        const newToken = await this.getToken();
        
        const retryResponse = await fetch(`${N8N_BASE_URL}${endpoint}`, {
          ...options,
          headers: {
            ...options.headers,
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${newToken}`
          }
        });
        
        return retryResponse;
      }

      return response;
    } catch (error) {
      console.error('Authenticated request failed:', error);
      throw error;
    }
  }
}

// Singleton instance
export const authService = new AuthService();