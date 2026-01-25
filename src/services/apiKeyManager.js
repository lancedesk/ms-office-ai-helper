// API Key Manager
// Handles storage, retrieval, and validation of API keys for multiple providers

class APIKeyManager {
  constructor() {
    this.groqStorageKey = 'groq_api_key';
    this.geminiStorageKey = 'gemini_api_key';
    this.providerKey = 'ai_provider'; // Which provider is selected
  }

  /**
   * Save Groq API key to storage
   * @param {string} apiKey - The Groq API key to save
   * @returns {Promise<boolean>} Success status
   */
  async saveGroqApiKey(apiKey) {
    return this.saveKey(this.groqStorageKey, apiKey);
  }

  /**
   * Save Gemini API key to storage
   * @param {string} apiKey - The Gemini API key to save
   * @returns {Promise<boolean>} Success status
   */
  async saveGeminiApiKey(apiKey) {
    return this.saveKey(this.geminiStorageKey, apiKey);
  }

  /**
   * Generic key saving function
   * @private
   */
  async saveKey(storageKey, apiKey) {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
        Office.context.document.settings.set(storageKey, apiKey);
        await Office.context.document.settings.saveAsync();
      } else if (typeof localStorage !== 'undefined') {
        localStorage.setItem(storageKey, apiKey);
      }
      return true;
    } catch (error) {
      console.error(`Failed to save API key (${storageKey}):`, error);
      return false;
    }
  }

  /**
   * Get Groq API key from storage
   * @returns {Promise<string|null>} The API key or null if not found
   */
  async getGroqApiKey() {
    return this.getKey(this.groqStorageKey);
  }

  /**
   * Get Gemini API key from storage
   * @returns {Promise<string|null>} The API key or null if not found
   */
  async getGeminiApiKey() {
    return this.getKey(this.geminiStorageKey);
  }

  /**
   * Generic key retrieval function
   * @private
   */
  async getKey(storageKey) {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
        return Office.context.document.settings.get(storageKey) || null;
      } else if (typeof localStorage !== 'undefined') {
        return localStorage.getItem(storageKey);
      } else {
        return null;
      }
    } catch (error) {
      console.error(`Failed to get API key (${storageKey}):`, error);
      return null;
    }
  }

  /**
   * Set the active AI provider
   * @param {string} provider - 'groq' or 'gemini'
   * @returns {Promise<boolean>} Success status
   */
  async setActiveProvider(provider) {
    if (!['groq', 'gemini'].includes(provider)) {
      throw new Error('Invalid provider. Must be "groq" or "gemini"');
    }
    return this.saveKey(this.providerKey, provider);
  }

  /**
   * Get the active AI provider
   * @returns {Promise<string>} 'groq' or 'gemini' (defaults to 'groq')
   */
  async getActiveProvider() {
    const provider = await this.getKey(this.providerKey);
    return provider || 'groq'; // Default to Groq
  }

  /**
   * Delete a specific API key
   * @param {string} provider - 'groq' or 'gemini'
   * @returns {Promise<boolean>} Success status
   */
  async deleteApiKey(provider) {
    try {
      const storageKey = provider === 'groq' ? this.groqStorageKey : this.geminiStorageKey;
      
      if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
        Office.context.document.settings.remove(storageKey);
        await Office.context.document.settings.saveAsync();
      } else if (typeof localStorage !== 'undefined') {
        localStorage.removeItem(storageKey);
      }
      return true;
    } catch (error) {
      console.error(`Failed to delete API key for ${provider}:`, error);
      return false;
    }
  }

  /**
   * Check if API key exists
   * @param {string} provider - 'groq' or 'gemini' (optional, checks both if not specified)
   * @returns {Promise<boolean>} True if API key(s) exist
   */
  async hasApiKey(provider = null) {
    if (provider === 'groq') {
      const key = await this.getGroqApiKey();
      return key !== null && key.trim().length > 0;
    } else if (provider === 'gemini') {
      const key = await this.getGeminiApiKey();
      return key !== null && key.trim().length > 0;
    } else {
      // Check if either provider has a key
      const groq = await this.hasApiKey('groq');
      const gemini = await this.hasApiKey('gemini');
      return groq || gemini;
    }
  }

  /**
   * Validate API key format
   * @param {string} key - The API key to validate
   * @param {string} provider - 'groq' or 'gemini'
   * @returns {boolean} True if format appears valid
   */
  validateFormat(key, provider) {
    if (!key || typeof key !== 'string') {
      return false;
    }

    if (provider === 'groq') {
      // Groq API keys start with 'gsk_' and are longer than 20 chars
      return key.startsWith('gsk_') && key.length > 20;
    } else if (provider === 'gemini') {
      // Google API keys start with 'AIza' and are longer than 30 chars
      return key.startsWith('AIza') && key.length > 30;
    }

    return false;
  }

  /**
   * Mask API key for display
   * @param {string} key - The API key to mask
   * @param {string} provider - 'groq' or 'gemini'
   * @returns {string} Masked API key
   */
  maskApiKey(key, provider) {
    if (!key || key.length < 10) {
      return '****';
    }
    return `${key.substring(0, 7)}...${key.substring(key.length - 4)}`;
  }

  /**
   * Get provider display name
   * @param {string} provider - 'groq' or 'gemini'
   * @returns {string} Display name
   */
  getProviderName(provider) {
    if (provider === 'groq') return 'Groq (Llama 3.1)';
    if (provider === 'gemini') return 'Google Gemini';
    return 'Unknown';
  }

  /**
   * Get provider description
   * @param {string} provider - 'groq' or 'gemini'
   * @returns {string} Description
   */
  getProviderDescription(provider) {
    if (provider === 'groq') {
      return 'Fast and capable open-source models via Groq';
    }
    if (provider === 'gemini') {
      return 'Google\'s advanced AI model with multimodal capabilities';
    }
    return '';
  }
}

// Export for use in other files
module.exports = APIKeyManager;

