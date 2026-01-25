// Groq API Service
// Handles all communication with Groq AI API

class GroqService {
  constructor() {
    this.apiKey = null;
    this.baseUrl = 'https://api.groq.com/openai/v1/chat/completions';
    this.model = 'llama-3.1-8b-instant'; // Updated default model (supported)
  }

  /**
   * Set the API key for Groq
   * @param {string} key - The Groq API key
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
    // Groq API keys start with 'gsk_'
    return key && key.startsWith('gsk_') && key.length > 20;
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
   * Send a chat message to Groq API
   * @param {Array} messages - Array of message objects with role and content
   * @param {Object} options - Optional parameters (temperature, max_tokens, etc.)
   * @returns {Promise<string>} The AI response text
   */
  async sendMessage(messages, options = {}) {
    if (!this.hasApiKey()) {
      throw new Error('API key not configured. Please set up your Groq API key first.');
    }

    const requestBody = {
      model: options.model || this.model,
      messages: messages,
      temperature: options.temperature || 0.7,
      max_tokens: options.max_tokens || 2000,
      top_p: options.top_p || 1,
      stream: false
    };

    try {
      const response = await fetch(this.baseUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${this.apiKey}`
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(
          errorData.error?.message || 
          `API request failed: ${response.status} ${response.statusText}`
        );
      }

      const data = await response.json();
      
      if (!data.choices || data.choices.length === 0) {
        throw new Error('No response from API');
      }

      return data.choices[0].message.content;
    } catch (error) {
      console.error('Groq API Error:', error);
      
      // Provide user-friendly error messages
      if (error.message.includes('401') || error.message.includes('Unauthorized')) {
        throw new Error('Invalid API key. Please check your Groq API key.');
      } else if (error.message.includes('429') || error.message.includes('rate limit')) {
        throw new Error('Rate limit exceeded. Please wait a moment and try again.');
      } else if (error.message.includes('network') || error.message.includes('Failed to fetch')) {
        throw new Error('Network error. Please check your internet connection.');
      }
      
      throw error;
    }
  }

  /**
   * Send a chat with streaming response (for future implementation)
   * @param {Array} messages - Array of message objects
   * @param {Function} onChunk - Callback for each chunk of response
   * @param {Object} options - Optional parameters
   */
  async sendMessageStream(messages, onChunk, options = {}) {
    if (!this.hasApiKey()) {
      throw new Error('API key not configured');
    }

    const requestBody = {
      model: options.model || this.model,
      messages: messages,
      temperature: options.temperature || 0.7,
      max_tokens: options.max_tokens || 2000,
      stream: true
    };

    try {
      const response = await fetch(this.baseUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${this.apiKey}`
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        throw new Error(`API request failed: ${response.status}`);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop(); // Keep incomplete line in buffer

        for (const line of lines) {
          if (line.startsWith('data: ')) {
            const data = line.slice(6);
            if (data === '[DONE]') continue;

            try {
              const json = JSON.parse(data);
              const content = json.choices?.[0]?.delta?.content;
              if (content) {
                onChunk(content);
              }
            } catch (e) {
              // Skip invalid JSON
            }
          }
        }
      }
    } catch (error) {
      console.error('Groq Streaming Error:', error);
      throw error;
    }
  }

  /**
   * Get available models
   * @returns {Array} List of available Groq models
   */
  getAvailableModels() {
    return [
      { id: 'llama-3.1-70b-versatile', name: 'Llama 3.1 70B (Fast & Versatile)' },
      { id: 'llama-3.1-8b-instant', name: 'Llama 3.1 8B (Instant)' },
      { id: 'mixtral-8x7b-32768', name: 'Mixtral 8x7B (Large Context)' },
      { id: 'gemma-7b-it', name: 'Gemma 7B' }
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
if (typeof module !== 'undefined' && module.exports) {
  module.exports = GroqService;
}
