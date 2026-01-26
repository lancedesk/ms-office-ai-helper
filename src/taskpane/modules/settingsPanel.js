// Settings Panel Module
// Handles API key setup, provider switching, and settings UI

import { removeWelcomeScreen, addSystemMessage, addAssistantMessage, scrollToBottom } from './chatUI.js';

// Track if settings panel is open
let isSettingsOpen = false;

// These will be set by the main module
let chatContainer = null;
let apiKeyManager = null;
let groqService = null;
let geminiService = null;
let currentProvider = 'groq';

/**
 * Initialize settings panel with required dependencies
 */
function initSettingsPanel(deps) {
  chatContainer = deps.chatContainer;
  apiKeyManager = deps.apiKeyManager;
  groqService = deps.groqService;
  geminiService = deps.geminiService;
  currentProvider = deps.currentProvider || 'groq';
}

/**
 * Update the current provider reference
 */
function setCurrentProvider(provider) {
  currentProvider = provider;
}

/**
 * Get current provider
 */
function getCurrentProvider() {
  return currentProvider;
}

/**
 * Check if settings is open
 */
function getIsSettingsOpen() {
  return isSettingsOpen;
}

/**
 * Toggle settings panel open/closed
 */
function toggleSettings() {
  if (isSettingsOpen) {
    closeSettingsPanel();
  } else {
    showApiKeySettings();
  }
}

/**
 * Show API key settings - either setup or panel based on existing keys
 */
async function showApiKeySettings() {
  var hasGroqKey = await apiKeyManager.hasApiKey('groq');
  var hasGeminiKey = await apiKeyManager.hasApiKey('gemini');
  
  if (hasGroqKey || hasGeminiKey) {
    showSettingsPanel(hasGroqKey, hasGeminiKey);
  } else {
    showApiKeySetup();
  }
}

/**
 * Show initial API key setup UI
 */
function showApiKeySetup() {
  removeWelcomeScreen();
  
  isSettingsOpen = true;
  
  var setupDiv = document.createElement("div");
  setupDiv.id = "api-key-setup";
  setupDiv.className = "api-key-setup";
  setupDiv.innerHTML = 
    '<div class="setup-content">' +
      '<h2>🔑 Welcome to AI Helper!</h2>' +
      '<p>Choose an AI provider and get your API key to get started.</p>' +
      
      '<div class="provider-tabs">' +
        '<button class="provider-tab active" onclick="switchSetupProvider(\'groq\')">' +
          '⚡ Groq (Llama 3.1)' +
        '</button>' +
        '<button class="provider-tab" onclick="switchSetupProvider(\'gemini\')">' +
          '🧠 Google Gemini' +
        '</button>' +
      '</div>' +

      '<div id="groq-setup" class="provider-setup active">' +
        '<div class="setup-steps">' +
          '<div class="setup-step">' +
            '<strong>Step 1:</strong> Create a free account at Groq<br>' +
            '<button onclick="window.open(\'https://console.groq.com\', \'_blank\')" class="link-button">' +
              'Open Groq Console →' +
            '</button>' +
          '</div>' +
          '<div class="setup-step"><strong>Step 2:</strong> Navigate to API Keys section</div>' +
          '<div class="setup-step"><strong>Step 3:</strong> Create a new API key and copy it</div>' +
          '<div class="setup-step">' +
            '<strong>Step 4:</strong> Paste your API key below' +
            '<input type="password" id="groq-api-key-input" placeholder="gsk_..." class="api-key-input" />' +
            '<button id="groq-show-key-btn" class="show-key-btn" onclick="toggleApiKeyVisibility(\'groq\')">👁️ Show</button>' +
          '</div>' +
        '</div>' +
        '<div id="groq-api-key-error" class="api-key-error"></div>' +
        '<div class="setup-buttons">' +
          '<button id="groq-test-key-btn" class="secondary-button">Test Connection</button>' +
          '<button id="groq-save-key-btn" class="primary-button">Save Groq Key</button>' +
        '</div>' +
      '</div>' +

      '<div id="gemini-setup" class="provider-setup">' +
        '<div class="setup-steps">' +
          '<div class="setup-step">' +
            '<strong>Step 1:</strong> Go to Google AI Studio<br>' +
            '<button onclick="window.open(\'https://aistudio.google.com/apikey\', \'_blank\')" class="link-button">' +
              'Open Google AI Studio →' +
            '</button>' +
          '</div>' +
          '<div class="setup-step"><strong>Step 2:</strong> Create a new API key (or use existing)</div>' +
          '<div class="setup-step"><strong>Step 3:</strong> Copy your API key from Google</div>' +
          '<div class="setup-step">' +
            '<strong>Step 4:</strong> Paste your API key below' +
            '<input type="password" id="gemini-api-key-input" placeholder="AIza..." class="api-key-input" />' +
            '<button id="gemini-show-key-btn" class="show-key-btn" onclick="toggleApiKeyVisibility(\'gemini\')">👁️ Show</button>' +
          '</div>' +
        '</div>' +
        '<div id="gemini-api-key-error" class="api-key-error"></div>' +
        '<div class="setup-buttons">' +
          '<button id="gemini-test-key-btn" class="secondary-button">Test Connection</button>' +
          '<button id="gemini-save-key-btn" class="primary-button">Save Gemini Key</button>' +
        '</div>' +
      '</div>' +
      
      '<p class="privacy-note">🔒 Your API keys are stored securely and never shared.</p>' +
    '</div>';
  
  chatContainer.appendChild(setupDiv);
  
  // Add event listeners
  document.getElementById("groq-test-key-btn").onclick = function() { testApiKey('groq'); };
  document.getElementById("groq-save-key-btn").onclick = function() { saveApiKey('groq'); };
  document.getElementById("gemini-test-key-btn").onclick = function() { testApiKey('gemini'); };
  document.getElementById("gemini-save-key-btn").onclick = function() { saveApiKey('gemini'); };
  
  document.getElementById("groq-api-key-input").focus();
}

/**
 * Switch between provider tabs in setup
 */
function switchSetupProvider(provider) {
  document.getElementById('groq-setup').classList.remove('active');
  document.getElementById('gemini-setup').classList.remove('active');
  
  document.querySelectorAll('.provider-tab').forEach(function(tab) {
    tab.classList.remove('active');
  });
  
  document.getElementById(provider + '-setup').classList.add('active');
  
  if (event && event.target) {
    event.target.classList.add('active');
  }
  
  var input = document.getElementById(provider + '-api-key-input');
  if (input) input.focus();
}

/**
 * Toggle API key visibility (show/hide)
 */
function toggleApiKeyVisibility(provider) {
  var input = document.getElementById(provider + '-api-key-input');
  var btn = document.getElementById(provider + '-show-key-btn');
  if (input.type === "password") {
    input.type = "text";
    btn.textContent = "🙈 Hide";
  } else {
    input.type = "password";
    btn.textContent = "👁️ Show";
  }
}

/**
 * Test an API key
 */
async function testApiKey(provider) {
  var input = document.getElementById(provider + '-api-key-input');
  var errorDiv = document.getElementById(provider + '-api-key-error');
  var testBtn = document.getElementById(provider + '-test-key-btn');
  var apiKey = input.value.trim();
  
  errorDiv.textContent = "";
  errorDiv.className = "api-key-error";
  
  if (!apiKey) {
    errorDiv.textContent = "⚠️ Please enter an API key";
    errorDiv.className = "api-key-error error";
    return;
  }
  
  if (!apiKeyManager.validateFormat(apiKey, provider)) {
    var format = provider === 'groq' ? "gsk_" : "AIza";
    errorDiv.textContent = "⚠️ Invalid API key format. " + (provider === 'groq' ? 'Groq' : 'Google') + " keys start with '" + format + "'";
    errorDiv.className = "api-key-error error";
    return;
  }
  
  testBtn.disabled = true;
  testBtn.textContent = "Testing...";
  
  try {
    var result;
    if (provider === 'groq') {
      groqService.setApiKey(apiKey);
      result = await groqService.testApiKey();
    } else {
      geminiService.setApiKey(apiKey);
      result = await geminiService.testApiKey();
    }
    
    if (result.valid) {
      errorDiv.textContent = "✅ API key is valid!";
      errorDiv.className = "api-key-error success";
    } else {
      errorDiv.textContent = "❌ " + result.error;
      errorDiv.className = "api-key-error error";
    }
  } catch (error) {
    errorDiv.textContent = "❌ Error: " + error.message;
    errorDiv.className = "api-key-error error";
  } finally {
    testBtn.disabled = false;
    testBtn.textContent = "Test Connection";
  }
}

/**
 * Save an API key from setup
 */
async function saveApiKey(provider) {
  var input = document.getElementById(provider + '-api-key-input');
  var errorDiv = document.getElementById(provider + '-api-key-error');
  var saveBtn = document.getElementById(provider + '-save-key-btn');
  var apiKey = input.value.trim();
  
  errorDiv.textContent = "";
  errorDiv.className = "api-key-error";
  
  if (!apiKey) {
    errorDiv.textContent = "⚠️ Please enter an API key";
    errorDiv.className = "api-key-error error";
    return;
  }
  
  if (!apiKeyManager.validateFormat(apiKey, provider)) {
    errorDiv.textContent = "⚠️ Invalid API key format";
    errorDiv.className = "api-key-error error";
    return;
  }
  
  saveBtn.disabled = true;
  saveBtn.textContent = "Saving...";
  
  try {
    var saved;
    if (provider === 'groq') {
      saved = await apiKeyManager.saveGroqApiKey(apiKey);
      groqService.setApiKey(apiKey);
    } else {
      saved = await apiKeyManager.saveGeminiApiKey(apiKey);
      geminiService.setApiKey(apiKey);
    }
    
    if (saved) {
      var hasOtherKey = provider === 'groq' 
        ? await apiKeyManager.hasApiKey('gemini')
        : await apiKeyManager.hasApiKey('groq');
      
      if (!hasOtherKey) {
        currentProvider = provider;
        await apiKeyManager.setActiveProvider(provider);
      }
      
      var setupDiv = document.getElementById("api-key-setup");
      if (setupDiv) setupDiv.remove();
      
      addSystemMessage("✅ " + apiKeyManager.getProviderName(provider) + " key saved successfully!");
      addAssistantMessage("Great! I'm ready to chat using " + apiKeyManager.getProviderName(provider) + ".\n\nI can help you with:\n\n• Summarizing documents\n• Editing and formatting text\n• Answering questions about your document\n• Creating tables, headers, and more\n\nWhat would you like to do?");
    } else {
      errorDiv.textContent = "❌ Failed to save " + (provider === 'groq' ? 'Groq' : 'Gemini') + " API key";
      errorDiv.className = "api-key-error error";
    }
  } catch (error) {
    errorDiv.textContent = "❌ Error: " + error.message;
    errorDiv.className = "api-key-error error";
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = "Save " + (provider === 'groq' ? 'Groq' : 'Gemini') + " Key";
  }
}

/**
 * Show settings panel (when keys already configured)
 */
async function showSettingsPanel(hasGroqKey, hasGeminiKey) {
  removeWelcomeScreen();
  
  var existingPanel = document.getElementById('settings-panel');
  if (existingPanel) existingPanel.remove();
  
  isSettingsOpen = true;
  
  var currentProviderName = apiKeyManager.getProviderName(currentProvider);
  
  var panelDiv = document.createElement('div');
  panelDiv.id = 'settings-panel';
  panelDiv.className = 'api-key-setup';
  panelDiv.innerHTML = 
    '<div class="setup-content">' +
      '<h2>⚙️ Settings</h2>' +
      
      '<div class="settings-section">' +
        '<h3 style="color: #667eea; margin-bottom: 12px;">🎯 Active Provider</h3>' +
        '<p style="color: #666; font-size: 13px; margin-bottom: 12px;">Currently using: <strong>' + currentProviderName + '</strong></p>' +
        
        '<div class="provider-selector" style="display: flex; gap: 10px; margin-bottom: 20px;">' +
          '<button id="select-groq-btn" class="' + (currentProvider === 'groq' ? 'primary-button' : 'secondary-button') + '" ' +
                  'style="flex: 1;' + (!hasGroqKey ? ' opacity: 0.5;' : '') + '" ' +
                  (hasGroqKey ? '' : 'disabled') + '>' +
            '⚡ Groq ' + (hasGroqKey ? '✓' : '(not set)') +
          '</button>' +
          '<button id="select-gemini-btn" class="' + (currentProvider === 'gemini' ? 'primary-button' : 'secondary-button') + '" ' +
                  'style="flex: 1;' + (!hasGeminiKey ? ' opacity: 0.5;' : '') + '" ' +
                  (hasGeminiKey ? '' : 'disabled') + '>' +
            '🧠 Gemini ' + (hasGeminiKey ? '✓' : '(not set)') +
          '</button>' +
        '</div>' +
      '</div>' +
      
      '<div class="settings-section" style="border-top: 1px solid #eee; padding-top: 20px;">' +
        '<h3 style="color: #667eea; margin-bottom: 12px;">🔑 API Keys</h3>' +
        
        '<div class="setup-step" style="margin-bottom: 12px;">' +
          '<strong>Groq API Key ' + (hasGroqKey ? '✅' : '') + '</strong>' +
          '<div style="display: flex; gap: 8px; margin-top: 8px;">' +
            '<input type="password" id="settings-groq-key" placeholder="' + (hasGroqKey ? '••••••••••••••••' : 'gsk_...') + '" class="api-key-input" style="flex: 1;" />' +
            '<button id="save-groq-key-btn" class="secondary-button" style="flex: none; padding: 10px 16px;">Save</button>' +
          '</div>' +
        '</div>' +
        
        '<div class="setup-step">' +
          '<strong>Gemini API Key ' + (hasGeminiKey ? '✅' : '') + '</strong>' +
          '<div style="display: flex; gap: 8px; margin-top: 8px;">' +
            '<input type="password" id="settings-gemini-key" placeholder="' + (hasGeminiKey ? '••••••••••••••••' : 'AIza...') + '" class="api-key-input" style="flex: 1;" />' +
            '<button id="save-gemini-key-btn" class="secondary-button" style="flex: none; padding: 10px 16px;">Save</button>' +
          '</div>' +
        '</div>' +
        
        '<div id="settings-status" class="api-key-error" style="margin-top: 12px;"></div>' +
      '</div>' +
      
      '<div class="setup-buttons" style="margin-top: 20px;">' +
        '<button id="close-settings-btn" class="primary-button">Close Settings</button>' +
      '</div>' +
      
      '<p class="privacy-note">🔒 Your API keys are stored locally and never shared.</p>' +
    '</div>';
  
  chatContainer.appendChild(panelDiv);
  scrollToBottom();
  
  // Attach event listeners
  document.getElementById('select-groq-btn').onclick = function() { switchActiveProvider('groq'); };
  document.getElementById('select-gemini-btn').onclick = function() { switchActiveProvider('gemini'); };
  document.getElementById('save-groq-key-btn').onclick = function() { saveKeyFromSettings('groq'); };
  document.getElementById('save-gemini-key-btn').onclick = function() { saveKeyFromSettings('gemini'); };
  document.getElementById('close-settings-btn').onclick = closeSettingsPanel;
}

/**
 * Switch the active AI provider
 */
async function switchActiveProvider(provider) {
  var hasKey = await apiKeyManager.hasApiKey(provider);
  if (!hasKey) {
    var statusDiv = document.getElementById('settings-status');
    statusDiv.textContent = "⚠️ Please add a " + apiKeyManager.getProviderName(provider) + " API key first";
    statusDiv.className = 'api-key-error error';
    return;
  }
  
  currentProvider = provider;
  await apiKeyManager.setActiveProvider(provider);
  
  var groqBtn = document.getElementById('select-groq-btn');
  var geminiBtn = document.getElementById('select-gemini-btn');
  
  if (provider === 'groq') {
    groqBtn.className = 'primary-button';
    geminiBtn.className = 'secondary-button';
  } else {
    groqBtn.className = 'secondary-button';
    geminiBtn.className = 'primary-button';
  }
  
  var statusDiv = document.getElementById('settings-status');
  statusDiv.textContent = "✅ Switched to " + apiKeyManager.getProviderName(provider);
  statusDiv.className = 'api-key-error success';
}

/**
 * Save API key from settings panel
 */
async function saveKeyFromSettings(provider) {
  var inputId = provider === 'groq' ? 'settings-groq-key' : 'settings-gemini-key';
  var input = document.getElementById(inputId);
  var statusDiv = document.getElementById('settings-status');
  var apiKey = input.value.trim();
  
  if (!apiKey) {
    statusDiv.textContent = '⚠️ Please enter an API key';
    statusDiv.className = 'api-key-error error';
    return;
  }
  
  if (!apiKeyManager.validateFormat(apiKey, provider)) {
    var format = provider === 'groq' ? 'gsk_' : 'AIza';
    statusDiv.textContent = "⚠️ Invalid format. " + apiKeyManager.getProviderName(provider) + " keys start with '" + format + "'";
    statusDiv.className = 'api-key-error error';
    return;
  }
  
  try {
    if (provider === 'groq') {
      await apiKeyManager.saveGroqApiKey(apiKey);
      groqService.setApiKey(apiKey);
    } else {
      await apiKeyManager.saveGeminiApiKey(apiKey);
      geminiService.setApiKey(apiKey);
    }
    
    statusDiv.textContent = "✅ " + apiKeyManager.getProviderName(provider) + " key saved!";
    statusDiv.className = 'api-key-error success';
    input.value = '';
    input.placeholder = '••••••••••••••••';
    
    var btnId = provider === 'groq' ? 'select-groq-btn' : 'select-gemini-btn';
    var btn = document.getElementById(btnId);
    btn.disabled = false;
    btn.style.opacity = '1';
    btn.innerHTML = provider === 'groq' ? '⚡ Groq ✓' : '🧠 Gemini ✓';
    
  } catch (error) {
    statusDiv.textContent = "❌ Error: " + error.message;
    statusDiv.className = 'api-key-error error';
  }
}

/**
 * Close settings panel
 */
function closeSettingsPanel() {
  var panel = document.getElementById('settings-panel');
  if (panel) panel.remove();
  
  var setupPanel = document.getElementById('api-key-setup');
  if (setupPanel) setupPanel.remove();
  
  isSettingsOpen = false;
}

export {
  initSettingsPanel,
  setCurrentProvider,
  getCurrentProvider,
  getIsSettingsOpen,
  toggleSettings,
  showApiKeySettings,
  showApiKeySetup,
  switchSetupProvider,
  toggleApiKeyVisibility,
  testApiKey,
  saveApiKey,
  showSettingsPanel,
  switchActiveProvider,
  saveKeyFromSettings,
  closeSettingsPanel
};
