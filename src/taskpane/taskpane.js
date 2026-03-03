/* global Office, Word */

// Import services (will be bundled by webpack)
import GroqService from '../services/groqService';
import GeminiService from '../services/geminiService';
import APIKeyManager from '../services/apiKeyManager';
import DocumentService from '../services/documentService';

// Import modules
import {
  initializeElements,
  getChatContainer,
  getMessageInput,
  getSendButton,
  removeWelcomeScreen,
  addSystemMessage,
  addUserMessage,
  addAssistantMessage,
  showLoading,
  hideLoading,
  scrollToBottom
} from './modules/chatUI.js';

import {
  shouldIncludeDocumentContext,
  buildSystemContext
} from './modules/aiContext.js';

import {
  parseAndExecuteActions,
  showReplaceConfirmation
} from './modules/actionExecutor.js';

import {
  initSettingsPanel,
  setCurrentProvider,
  getCurrentProvider,
  toggleSettings,
  showApiKeySettings,
  switchSetupProvider,
  toggleApiKeyVisibility,
  switchActiveProvider,
  saveKeyFromSettings,
  closeSettingsPanel
} from './modules/settingsPanel.js';

import {
  initSpecialCommands,
  handleSpecialCommand
} from './modules/specialCommands.js';

// Wrap everything in try-catch to catch any initialization errors
try {

// Debug logging helper - console only (no alert, no DOM manipulation at top level)
function debugLog(message) {
  console.log('[AI Helper] ' + message);
}

debugLog('Bundle loaded, initializing services...');

// Initialize services
var groqService = new GroqService();
var geminiService = new GeminiService();
var apiKeyManager = new APIKeyManager();
var documentService = new DocumentService();

debugLog('Services initialized successfully');

// AI Provider (can be 'groq' or 'gemini')
var currentProvider = 'groq';

// Chat history
var chatHistory = [];

// Track if we've already initialized
var isInitialized = false;

// Safe initialization function
function safeInitialize() {
  if (isInitialized) {
    debugLog('Already initialized, skipping');
    return;
  }
  
  try {
    debugLog('Starting initialization...');
    
    // Initialize DOM elements
    initializeElements();
    
    // Initialize modules with dependencies
    initSettingsPanel({
      chatContainer: getChatContainer(),
      apiKeyManager: apiKeyManager,
      groqService: groqService,
      geminiService: geminiService,
      currentProvider: currentProvider
    });
    
    initSpecialCommands({
      documentService: documentService
    });
    
    // Attach event handlers
    attachEventHandlers();
    
    // Initialize app
    initializeApp();
    
    isInitialized = true;
    debugLog('Initialization complete!');
  } catch (err) {
    debugLog('Init error: ' + err.message);
    console.error('Initialization error:', err);
  }
}

// Try Office.onReady first
if (typeof Office !== 'undefined' && Office.onReady) {
  debugLog('Office.js detected, waiting for onReady...');
  Office.onReady(function(info) {
    debugLog('Office.onReady fired. Host: ' + (info ? info.host : 'unknown'));
    safeInitialize();
  });
} else {
  debugLog('Office.js not detected');
}

// Fallback: Initialize on DOMContentLoaded if Office.onReady doesn't fire
document.addEventListener('DOMContentLoaded', function() {
  debugLog('DOMContentLoaded fired');
  // Give Office.onReady a chance to fire first (500ms)
  setTimeout(function() {
    if (!isInitialized) {
      debugLog('Fallback init after timeout');
      safeInitialize();
    }
  }, 500);
});

// Extra fallback: window.onload
window.addEventListener('load', function() {
  debugLog('window.load fired');
  setTimeout(function() {
    if (!isInitialized) {
      debugLog('window.load fallback init');
      safeInitialize();
    }
  }, 1000);
});

/**
 * Attach UI event handlers safely (idempotent)
 */
function attachEventHandlers() {
  try {
    var sendButton = getSendButton();
    var messageInput = getMessageInput();

    if (sendButton) {
      sendButton.onclick = null;
      sendButton.addEventListener('click', sendMessage);
    }

    var settingsBtn = document.getElementById('settings-button');
    if (settingsBtn) {
      settingsBtn.onclick = null;
      settingsBtn.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        toggleSettings();
      });
    }

    if (messageInput) {
      messageInput.onkeypress = null;
      messageInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          sendMessage();
        }
      });
      messageInput.addEventListener('input', function() {
        messageInput.style.height = 'auto';
        messageInput.style.height = Math.max(44, Math.min(messageInput.scrollHeight, 150)) + 'px';
      });
    }
  } catch (err) {
    console.warn('attachEventHandlers error:', err);
  }
}

async function initializeApp() {
  try {
    var savedGroqKey = await apiKeyManager.getGroqApiKey();
    var savedGeminiKey = await apiKeyManager.getGeminiApiKey();
    var savedProvider = await apiKeyManager.getActiveProvider();
    
    if (savedGroqKey) {
      groqService.setApiKey(savedGroqKey);
    }
    if (savedGeminiKey) {
      geminiService.setApiKey(savedGeminiKey);
    }
    
    currentProvider = savedProvider || 'groq';
    setCurrentProvider(currentProvider);
    
    var currentService = currentProvider === 'groq' ? groqService : geminiService;
    if (currentService.hasApiKey()) {
      // Ready to chat - no banner message
    } else {
      addSystemMessage("👋 Welcome! Please configure your API key to get started.");
      showApiKeySettings();
    }
  } catch (err) {
    console.warn('Error loading saved keys:', err);
    addSystemMessage("👋 Welcome! Please configure your API key to get started.");
    showApiKeySettings();
  }
}

async function sendMessage() {
  var messageInput = getMessageInput();
  var sendButton = getSendButton();

  var message = messageInput.value.trim();
  if (!message) return;

  // Check for special commands first (starting with /)
  if (message.charAt(0) === '/') {
    addUserMessage(message);
    messageInput.value = "";
    var handled = await handleSpecialCommand(message);
    if (handled) return;
  }
  
  // Show user message first
  addUserMessage(message);
  messageInput.value = "";
  messageInput.style.height = '44px';
  
  // Check if user is confirming a pending CREATE fallback
  if (window._pendingCreateContent && message.toLowerCase().match(/^(yes|yeah|ok|confirm|do it|replace)/)) {
    showLoading();
    try {
      var confirmed = await showReplaceConfirmation(window._pendingCreateContent);
      if (confirmed) {
        await documentService.replaceDocumentContent(window._pendingCreateContent);
        addSystemMessage("📝 Document has been replaced with the new content!");
      } else {
        addSystemMessage("❌ Document replacement cancelled.");
      }
    } catch (error) {
      addSystemMessage("⚠️ Failed to replace document: " + error.message);
    }
    window._pendingCreateContent = null;
    hideLoading();
    return;
  }
  
  // All requests go through AI - it interprets intent and returns structured actions

  // Check if API key is configured for selected provider
  var hasGroqKey = await apiKeyManager.hasApiKey('groq');
  var hasGeminiKey = await apiKeyManager.hasApiKey('gemini');

  if (!hasGroqKey && !hasGeminiKey) {
    addSystemMessage("❌ No API key configured. Please click the settings (⚙️) icon and add at least one provider's API key to use chat features.");
    messageInput.disabled = false;
    sendButton.disabled = false;
    return;
  }

  // If selected provider doesn't have a key, switch to available one
  currentProvider = getCurrentProvider();
  
  if (currentProvider === 'groq' && !hasGroqKey) {
    if (hasGeminiKey) {
      currentProvider = 'gemini';
      setCurrentProvider(currentProvider);
      addSystemMessage("⚠️ Switched to Gemini (Groq key not configured)");
    } else {
      addSystemMessage("❌ Please configure a Groq API key in settings.");
      messageInput.disabled = false;
      sendButton.disabled = false;
      return;
    }
  } else if (currentProvider === 'gemini' && !hasGeminiKey) {
    if (hasGroqKey) {
      currentProvider = 'groq';
      setCurrentProvider(currentProvider);
      addSystemMessage("⚠️ Switched to Groq (Gemini key not configured)");
    } else {
      addSystemMessage("❌ Please configure a Gemini API key in settings.");
      messageInput.disabled = false;
      sendButton.disabled = false;
      return;
    }
  }

  // Disable input while processing
  messageInput.disabled = true;
  sendButton.disabled = true;

  // Add to history
  chatHistory.push({ role: 'user', content: message });

  // Show loading indicator
  showLoading();

  try {
    // Determine if we need document context
    var needsDocContext = shouldIncludeDocumentContext(message);
    
    var systemContext = buildSystemContext();
    
    // Get document context if needed
    if (needsDocContext) {
      try {
        var docContext = await documentService.getDocumentContext();
        
        if (!docContext.isEmpty) {
          var formattedContext = documentService.formatContextForAI(docContext);
          systemContext += "\n\n" + formattedContext;
        } else {
          systemContext += "\n\nNote: The document is currently empty.";
        }
      } catch (error) {
        console.warn("Could not get document context:", error);
        systemContext += "\n\nNote: Unable to read document at this time.";
      }
    }

    // Prepare messages for API
    var messages = [
      { role: 'system', content: systemContext }
    ].concat(chatHistory.slice(-10));

    // Send to appropriate AI service
    var response;
    if (currentProvider === 'gemini') {
      response = await geminiService.sendMessage(messages);
    } else {
      response = await groqService.sendMessage(messages);
    }
    
    hideLoading();
    
    // Check for and execute any action commands in the response
    var processedResponse = await parseAndExecuteActions(response, documentService);
    
    // Add AI response to chat (with action tags removed)
    addAssistantMessage(processedResponse);
    
    // Add to chat history
    chatHistory.push({ role: 'assistant', content: response });
    
  } catch (error) {
    hideLoading();
    addAssistantMessage("❌ Error: " + error.message);
    console.error("Error:", error);
  } finally {
    // Re-enable input
    messageInput.disabled = false;
    sendButton.disabled = false;
    messageInput.focus();
  }
}

// Expose functions used by inline onclick attributes so they work in the browser and Word
if (typeof window !== 'undefined') {
  window.switchSetupProvider = switchSetupProvider;
  window.toggleApiKeyVisibility = toggleApiKeyVisibility;
  window.showApiKeySettings = showApiKeySettings;
  window.switchActiveProvider = switchActiveProvider;
  window.saveKeyFromSettings = saveKeyFromSettings;
  window.closeSettingsPanel = closeSettingsPanel;
}

} catch (bundleError) {
  // Catch any top-level errors and display them
  console.error('Bundle initialization error:', bundleError);
  var debugEl = document.getElementById('debug-output');
  if (debugEl) {
    debugEl.textContent = 'ERROR: ' + bundleError.message;
    debugEl.style.color = 'red';
    debugEl.style.background = '#ffe0e0';
  }
}
