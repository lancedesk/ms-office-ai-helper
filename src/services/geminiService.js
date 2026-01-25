// Google Gemini API Service
// Handles communication with Google's Gemini AI API

class GeminiService {
  constructor() {
    this.apiKey = null;
    this.baseUrl = 'https://generativelanguage.googleapis.com/v1beta/models';
    this.model = 'gemini-flash-latest'; // Use a model available to this API key
  }

  /**
   * Set the API key for Gemini
   * @param {string} key - The Google Gemini API key
   */
  setApiKey(key) {
    this.apiKey = key;
  }

  /**
   * Get the current API key
   * @returns {string|null} The API key or null if not set
   */
  getApiKey() {
    return this.apiKey;
  }

  /**
   * Check if API key is set
   * @returns {boolean} True if API key is configured
   */
  hasApiKey() {
    return this.apiKey !== null && this.apiKey.trim().length > 0;
  }

  /**
   * Validate API key format
   * @param {string} key - The API key to validate
   * @returns {boolean} True if key format is valid
   */
  isValidKeyFormat(key) {
    // Google API keys are usually long alphanumeric strings starting with 'AIza'
    return key && key.startsWith('AIza') && key.length > 30;
  }

  /**
   * Test API key by making a simple request
   * @returns {Promise<{valid: boolean, error?: string}>}
   */
  async testApiKey() {
    if (!this.hasApiKey()) {
      return { valid: false, error: 'No API key set' };
    }

    try {
      // sendMessage expects OpenAI-style messages: { role: 'user', content: '...'}
      const response = await this.sendMessage([
        { role: 'user', content: 'Hello' }
      ]);
      return { valid: true };
    } catch (error) {
      return { 
        valid: false, 
        error: error.message || 'Failed to validate API key' 
      };
    }
  }

  /**
   * Convert OpenAI format messages to Gemini format
   * @param {Array} messages - Messages in OpenAI format
   * @returns {Object} Gemini format conversation
   */
  convertMessagesToGemini(messages) {
    // Convert OpenAI-style messages to Gemini `contents` shape:
    // contents: [{ role: 'user'|'system'|'assistant', parts: [{ text: '...' }] }]
    const contents = [];
    for (const msg of messages) {
      const role = msg.role === 'user' ? 'user' : (msg.role === 'system' ? 'system' : 'assistant');
      contents.push({ role, parts: [{ text: String(msg.content) }] });
    }
    return { contents };
  }

  /**
   * Send a chat message to Gemini API
   * @param {Array} messages - Array of message objects
   * @param {Object} options - Optional parameters (temperature, max_tokens, etc.)
   * @returns {Promise<string>} The AI response text
   */
  async sendMessage(messages, options = {}) {
    if (!this.hasApiKey()) {
      throw new Error('API key not configured. Please set up your Gemini API key first.');
    }

    const { promptMessages } = this.convertMessagesToGemini(messages);

    // Build request body using the working `generateContent` shape.
    const { contents } = this.convertMessagesToGemini(messages);
    const requestBody = { contents };

    try {
      // Send the single, correct `generateContent` request
      const url = `${this.baseUrl}/${this.model}:generateContent?key=${this.apiKey}`;
      console.debug('Gemini generateContent URL:', url);
      console.debug('Gemini request body:', JSON.stringify(requestBody, null, 2));

      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody)
      });

      const data = await response.json().catch(() => null);

      if (!response.ok) {
        console.error('Gemini response error data:', data || await response.text());
        const msg = data?.error?.message || `HTTP ${response.status} ${response.statusText}`;
        throw new Error(msg);
      }

      // Extract text from canonical response shape
      if (data && data.candidates && data.candidates.length) {
        const content = data.candidates[0].content;
        if (content && content.parts && content.parts.length && content.parts[0].text) {
          return content.parts[0].text;
        }
      }

      throw new Error('No usable text returned from Gemini');
    } catch (error) {
      console.error('Gemini API Error:', error);
      // Provide helpful messages based on collected error text
      const msg = (error.message || '').toLowerCase();
      if (msg.includes('403') || msg.includes('api key') || msg.includes('invalid')) {
        throw new Error('Invalid API key or insufficient permissions for Gemini. Check your API key at https://aistudio.google.com/app/api-keys');
      }
      if (msg.includes('quota') || msg.includes('resource_exhausted')) {
        throw new Error('Quota exceeded or resource exhausted. Check your Google Cloud quotas or try again later.');
      }
      throw error;
    }
  }

  /**
   * Get available Gemini models
   * @returns {Array} List of available Gemini models
   */
  getAvailableModels() {
    return [
      { id: 'gemini-2.0-flash', name: 'Gemini 2.0 Flash (Latest & Fastest)' },
      { id: 'gemini-1.5-pro', name: 'Gemini 1.5 Pro (Most Capable)' },
      { id: 'gemini-1.5-flash', name: 'Gemini 1.5 Flash' },
      { id: 'gemini-1.0-pro', name: 'Gemini 1.0 Pro' }
    ];
  }

  /**
   * Change the model being used
   * @param {string} modelId - The model ID to use
   */
  setModel(modelId) {
    this.model = modelId;
  }
}

// Export for use in other files
module.exports = GeminiService;
