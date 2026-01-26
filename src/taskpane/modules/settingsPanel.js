// Settings Panel Module
// Handles API key configuration and settings UI

import { addSystemMessage, removeWelcomeScreen, scrollToBottom } from './chatUI.js';

// State
var isSettingsOpen = false;
var currentProvider = 'groq';

// Service references (set via initialize)
var apiKeyManager = null;
var groqService = null;
var geminiService = null;
var chatContainer = null;

/**
 * Initialize settings panel with required services
 */
function initializeSettings(services) {
  apiKeyManager = services.apiKeyManager;
  groqService = services.groqService;
  geminiService = services.geminiService;
  chatContainer = services.chatContainer;
  currentProvider = services.currentProvider || 'groq';
}

/**
 * Get current provider
 */
function getCurrentProvider() {
  return currentProvider;
}

/**
 * Set current provider
 */
function setCurrentProvider(provider) {
  currentProvider = provider;
}

/**
 * Check if settings panel is open
 */
function isSettingsVisible() {
  return isSettingsOpen;
}

/**
 * Toggle settings visibility
 */
async function toggleSettings() {
  if (isSettingsOpen) {
    closeSettingsPanel();
  } else {
    await showApiKeySettings();
  }
}

/**
 * Show API key settings
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
 * Show initial API key setup
 */
function showApiKeySetup() {
  removeWelcomeScreen();
  
  var existingSetup = document.getElementById('api-key-setup');
  if (existingSetup) {
    existingSetup.remove();
  }
  
  isSettingsOpen = true;
  
  var setupDiv = document.createElement('div');
  setupDiv.id = 'api-key-setup';
  setupDiv.className = 'api-key-setup';
  setupDiv.innerHTML = `
    <div class="setup-content">
      <h2>🔑 Welcome to AI Helper!</h2>
      <p>To get started, you'll need an API key from one of our supported providers:</p>
      
      <div class="provider-tabs">
        <button class="provider-tab active" data-provider="groq" onclick="switchSetupProvider('groq')">⚡ Groq (Fast)</button>
        <button class="provider-tab" data-provider="gemini" onclick="switchSetupProvider('gemini')">🧠 Gemini</button>
      </div>
      
      <div id="groq-setup" class="provider-setup active">
        <div class="setup-step">
          <strong>Step 1:</strong> Get your free Groq API key from 
          <a href="https://console.groq.com/keys" target="_blank" rel="noopener">console.groq.com</a>
        </div>
        <div class="setup-step">
          <strong>Step 2:</strong> Paste your API key below:
          <div class="input-group">
            <input type="password" id="groq-api-key" placeholder="gsk_..." class="api-key-input" />
            <button class="toggle-visibility" onclick="toggleApiKeyVisibility('groq')">👁️</button>
          </div>
        </div>
        <div class="setup-buttons">
          <button id="test-groq-key" class="secondary-button" onclick="testApiKey('groq')">Test Connection</button>
          <button id="save-groq-key" class="primary-button" onclick="saveApiKey('groq')">Save & Continue</button>
        </div>
      </div>
      
      <div id="gemini-setup" class="provider-setup">
        <div class="setup-step">
          <strong>Step 1:</strong> Get your free Gemini API key from 
          <a href="https://aistudio.google.com/apikey" target="_blank" rel="noopener">Google AI Studio</a>
        </div>
        <div class="setup-step">
          <strong>Step 2:</strong> Paste your API key below:
          <div class="input-group">
            <input type="password" id="gemini-api-key" placeholder="AIza..." class="api-key-input" />
            <button class="toggle-visibility" onclick="toggleApiKeyVisibility('gemini')">👁️</button>
          </div>
        </div>
        <div class="setup-buttons">
          <button id="test-gemini-key" class="secondary-button" onclick="testApiKey('gemini')">Test Connection</button>
          <button id="save-gemini-key" class="primary-button" onclick="saveApiKey('gemini')">Save & Continue</button>
        </div>
      </div>
      
      <div id="api-key-status" class="api-key-error"></div>
      
      <p class="privacy-note">🔒 Your API key is stored locally in your browser and is never sent to any server except the AI provider.</p>
    </div>
  `;
  
  chatContainer.appendChild(setupDiv);
  scrollToBottom();
}

/**
 * Switch between provider tabs in setup
 */
function switchSetupProvider(provider) {
  var tabs = document.querySelectorAll('.provider-tab');
  tabs.forEach(function(tab) {
    tab.classList.remove('active');
    if (tab.dataset.provider === provider) {
      tab.classList.add('active');
    }
  });
  
  var setups = document.querySelectorAll('.provider-setup');
  setups.forEach(function(setup) {
    setup.classList.remove('active');
  });
  
  document.getElementById(provider + '-setup').classList.add('active');
}

/**
 * Toggle API key visibility
 */
function toggleApiKeyVisibility(provider) {
  var input = document.getElementById(provider + '-api-key');
  if (input) {
    input.type = input.type === 'password' ? 'text' : 'password';
  }
}

/**
 * Test API key connection
 */
async function testApiKey(provider) {
  var input = document.getElementById(provider + '-api-key');
  var statusDiv = document.getElementById('api-key-status');
  var apiKey = input.value.trim();
  
  if (!apiKey) {
    statusDiv.textContent = '⚠️ Please enter an API key first';
    statusDiv.className = 'api-key-error error';
    return;
  }
  
  if (!apiKeyManager.validateFormat(apiKey, provider)) {
    var format = provider === 'groq' ? 'gsk_' : 'AIza';
    statusDiv.textContent = "⚠️ Invalid format. " + apiKeyManager.getProviderName(provider) + " keys start with '" + format + "'";
    statusDiv.className = 'api-key-error error';
    return;
  }
  
  statusDiv.textContent = '🔄 Testing connection...';
  statusDiv.className = 'api-key-error';
  
  try {
    var service = provider === 'groq' ? groqService : geminiService;
    service.setApiKey(apiKey);
    
    var testMessages = [
      { role: 'user', content: 'Say "Connection successful!" and nothing else.' }
    ];
    
    var response = await service.sendMessage(testMessages);
    
    if (response && response.length > 0) {
      statusDiv.textContent = '✅ Connection successful!';
      statusDiv.className = 'api-key-error success';
    } else {
      throw new Error('Empty response from API');
    }
  } catch (error) {
    statusDiv.textContent = '❌ Connection failed: ' + error.message;
    statusDiv.className = 'api-key-error error';
  }
}

/**
 * Save API key
 */
async function saveApiKey(provider) {
  var input = document.getElementById(provider + '-api-key');
  var statusDiv = document.getElementById('api-key-status');
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
    
    currentProvider = provider;
    await apiKeyManager.setActiveProvider(provider);
    
    statusDiv.textContent = '✅ API key saved successfully!';
    statusDiv.className = 'api-key-error success';
    
    setTimeout(function() {
      var setupPanel = document.getElementById('api-key-setup');
      if (setupPanel) {
        setupPanel.remove();
      }
      isSettingsOpen = false;
      addSystemMessage("🎉 You're all set! Start chatting by typing a message below.");
    }, 1000);
    
  } catch (error) {
    statusDiv.textContent = '❌ Error saving key: ' + error.message;
    statusDiv.className = 'api-key-error error';
  }
}

/**
 * Show settings panel
 */
async function showSettingsPanel(hasGroqKey, hasGeminiKey) {
  removeWelcomeScreen();
  
  var existingPanel = document.getElementById('settings-panel');
  if (existingPanel) {
    existingPanel.remove();
  }
  
  isSettingsOpen = true;
  
  var currentProviderName = apiKeyManager.getProviderName(currentProvider);
  
  var panelDiv = document.createElement('div');
  panelDiv.id = 'settings-panel';
  panelDiv.className = 'api-key-setup';
  panelDiv.innerHTML = `
    <div class="setup-content">
      <h2>⚙️ Settings</h2>
      
      <div class="settings-section">
        <h3 style="color: #667eea; margin-bottom: 12px;">🎯 Active Provider</h3>
        <p style="color: #666; font-size: 13px; margin-bottom: 12px;">Currently using: <strong>${currentProviderName}</strong></p>
        
        <div class="provider-selector" style="display: flex; gap: 10px; margin-bottom: 20px;">
          <button id="select-groq-btn" class="${currentProvider === 'groq' ? 'primary-button' : 'secondary-button'}" 
                  style="flex: 1; ${!hasGroqKey ? 'opacity: 0.5;' : ''}" 
                  ${!hasGroqKey ? 'disabled' : ''}>
            ⚡ Groq ${hasGroqKey ? '✓' : '(not set)'}
          </button>
          <button id="select-gemini-btn" class="${currentProvider === 'gemini' ? 'primary-button' : 'secondary-button'}" 
                  style="flex: 1; ${!hasGeminiKey ? 'opacity: 0.5;' : ''}" 
                  ${!hasGeminiKey ? 'disabled' : ''}>
            🧠 Gemini ${hasGeminiKey ? '✓' : '(not set)'}
          </button>
        </div>
      </div>
      
      <div class="settings-section" style="border-top: 1px solid #eee; padding-top: 20px;">
        <h3 style="color: #667eea; margin-bottom: 12px;">🔑 API Keys</h3>
        
        <div class="setup-step" style="margin-bottom: 12px;">
          <strong>Groq API Key ${hasGroqKey ? '✅' : ''}</strong>
          <div style="display: flex; gap: 8px; margin-top: 8px;">
            <input type="password" id="settings-groq-key" placeholder="${hasGroqKey ? '••••••••••••••••' : 'gsk_...'}" class="api-key-input" style="flex: 1;" />
            <button id="save-groq-key-btn" class="secondary-button" style="flex: none; padding: 10px 16px;">Save</button>
          </div>
        </div>
        
        <div class="setup-step">
          <strong>Gemini API Key ${hasGeminiKey ? '✅' : ''}</strong>
          <div style="display: flex; gap: 8px; margin-top: 8px;">
            <input type="password" id="settings-gemini-key" placeholder="${hasGeminiKey ? '••••••••••••••••' : 'AIza...'}" class="api-key-input" style="flex: 1;" />
            <button id="save-gemini-key-btn" class="secondary-button" style="flex: none; padding: 10px 16px;">Save</button>
          </div>
        </div>
        
        <div id="settings-status" class="api-key-error" style="margin-top: 12px;"></div>
      </div>
      
      <div class="setup-buttons" style="margin-top: 20px;">
        <button id="close-settings-btn" class="primary-button">Close Settings</button>
      </div>
      
      <p class="privacy-note">🔒 Your API keys are stored locally and never shared.</p>
    </div>
  `;
  
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
 * Switch active provider
 */
async function switchActiveProvider(provider) {
  var hasKey = await apiKeyManager.hasApiKey(provider);
  if (!hasKey) {
    var statusDiv = document.getElementById('settings-status');
    statusDiv.textContent = '⚠️ Please add a ' + apiKeyManager.getProviderName(provider) + ' API key first';
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
  statusDiv.textContent = '✅ Switched to ' + apiKeyManager.getProviderName(provider);
  statusDiv.className = 'api-key-error success';
}

/**
 * Save key from settings panel
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
    statusDiv.textContent = '⚠️ Invalid format. ' + apiKeyManager.getProviderName(provider) + " keys start with '" + format + "'";
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
    
    statusDiv.textContent = '✅ ' + apiKeyManager.getProviderName(provider) + ' key saved!';
    statusDiv.className = 'api-key-error success';
    input.value = '';
    input.placeholder = '••••••••••••••••';
    
    var btnId = provider === 'groq' ? 'select-groq-btn' : 'select-gemini-btn';
    var btn = document.getElementById(btnId);
    btn.disabled = false;
    btn.style.opacity = '1';
    btn.innerHTML = provider === 'groq' ? '⚡ Groq ✓' : '🧠 Gemini ✓';
    
  } catch (error) {
    statusDiv.textContent = '❌ Error: ' + error.message;
    statusDiv.className = 'api-key-error error';
  }
}

/**
 * Close settings panel
 */
function closeSettingsPanel() {
  var panel = document.getElementById('settings-panel');
  if (panel) {
    panel.remove();
  }
  var setupPanel = document.getElementById('api-key-setup');
  if (setupPanel) {
    setupPanel.remove();
  }
  isSettingsOpen = false;
}

// Export functions
export {
  initializeSettings,
  getCurrentProvider,
  setCurrentProvider,
  isSettingsVisible,
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
