/* global Office, Word */

// Import services (will be bundled by webpack)
import GroqService from '../services/groqService';
import GeminiService from '../services/geminiService';
import APIKeyManager from '../services/apiKeyManager';
import DocumentService from '../services/documentService';

// Wrap everything in try-catch to catch any initialization errors
try {

// Debug logging helper - console only (no alert, no DOM manipulation at top level)
function debugLog(message) {
  console.log('[AI Helper] ' + message);
}

debugLog('Bundle loaded, initializing services...');

// Initialize services
const groqService = new GroqService();
const geminiService = new GeminiService();
const apiKeyManager = new APIKeyManager();
const documentService = new DocumentService();

debugLog('Services initialized successfully');

// AI Provider (can be 'groq' or 'gemini')
let currentProvider = 'groq';

// Chat history
let chatHistory = [];

// Track if we've already initialized
let isInitialized = false;

// DOM element references - declared early to avoid temporal dead zone
let chatContainer = null;
let messageInput = null;
let sendButton = null;
let isWelcomeScreen = true;

// Safe initialization function
function safeInitialize() {
  if (isInitialized) {
    debugLog('Already initialized, skipping');
    return;
  }
  
  try {
    debugLog('Starting initialization...');
    attachEventHandlers();
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
  Office.onReady((info) => {
    debugLog('Office.onReady fired. Host: ' + (info ? info.host : 'unknown'));
    safeInitialize();
  });
} else {
  debugLog('Office.js not detected');
}

// Fallback: Initialize on DOMContentLoaded if Office.onReady doesn't fire
document.addEventListener('DOMContentLoaded', () => {
  debugLog('DOMContentLoaded fired');
  // Give Office.onReady a chance to fire first (500ms)
  setTimeout(() => {
    if (!isInitialized) {
      debugLog('Fallback init after timeout');
      safeInitialize();
    }
  }, 500);
});

// Extra fallback: window.onload
window.addEventListener('load', () => {
  debugLog('window.load fired');
  setTimeout(() => {
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
    // use initializeElements to populate references
    initializeElements();

    if (sendButton) {
      // remove previous listener if any
      sendButton.onclick = null;
      sendButton.addEventListener('click', sendMessage);
    }

    const settingsBtn = document.getElementById('settings-button');
    if (settingsBtn) {
      // ensure we don't attach multiple handlers
      settingsBtn.onclick = null;
      settingsBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        toggleSettings();
      });
    }

    if (messageInput) {
      // remove prior key handlers
      messageInput.onkeypress = null;
      messageInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
          sendMessage();
        }
      });
    }
  } catch (err) {
    console.warn('attachEventHandlers error:', err);
  }
}

async function initializeApp() {
  // Try to load saved API keys from storage
  try {
    const savedGroqKey = await apiKeyManager.getGroqApiKey();
    const savedGeminiKey = await apiKeyManager.getGeminiApiKey();
    const savedProvider = await apiKeyManager.getActiveProvider();
    
    if (savedGroqKey) {
      groqService.setApiKey(savedGroqKey);
    }
    if (savedGeminiKey) {
      geminiService.setApiKey(savedGeminiKey);
    }
    
    currentProvider = savedProvider || 'groq';
    
    // Check if we have a valid API key for the current provider
    const currentService = currentProvider === 'groq' ? groqService : geminiService;
    if (currentService.hasApiKey()) {
      addSystemMessage(`🎉 Ready to chat! Using: ${currentProvider === 'groq' ? 'Groq' : 'Google Gemini'}`);
    } else {
      // No API key configured - show setup
      addSystemMessage(`👋 Welcome! Please configure your API key to get started.`);
      showApiKeySettings();
    }
  } catch (err) {
    console.warn('Error loading saved keys:', err);
    addSystemMessage(`👋 Welcome! Please configure your API key to get started.`);
    showApiKeySettings();
  }
}

function initializeElements() {
  if (!chatContainer) {
    chatContainer = document.getElementById("chat-container");
    messageInput = document.getElementById("message-input");
    sendButton = document.getElementById("send-button");
  }
}

function removeWelcomeScreen() {
  if (isWelcomeScreen) {
    const welcomeScreen = document.querySelector(".welcome-screen");
    if (welcomeScreen) {
      welcomeScreen.remove();
    }
    isWelcomeScreen = false;
  }
}

function addSystemMessage(text) {
  initializeElements();
  removeWelcomeScreen();
  
  const messageDiv = document.createElement("div");
  messageDiv.className = "system-message";
  messageDiv.textContent = text;
  chatContainer.appendChild(messageDiv);
  scrollToBottom();
}

function addUserMessage(text) {
  initializeElements();
  removeWelcomeScreen();
  
  const messageDiv = document.createElement("div");
  messageDiv.className = "message user";
  messageDiv.innerHTML = `
    <div class="message-avatar">U</div>
    <div class="message-content">${escapeHtml(text)}</div>
  `;
  chatContainer.appendChild(messageDiv);
  scrollToBottom();
}

function addAssistantMessage(text) {
  initializeElements();
  
  var msgId = generateMessageId();
  var messageDiv = document.createElement("div");
  messageDiv.className = "message assistant";
  messageDiv.id = msgId;
  
  // Store raw text for copying
  messageDiv.setAttribute('data-raw-text', text);
  
  var contentHtml = parseMarkdown(text);
  
  messageDiv.innerHTML = 
    '<div class="message-avatar">AI</div>' +
    '<div class="message-content-wrapper">' +
      '<div class="message-content markdown-content">' + contentHtml + '</div>' +
      '<div class="message-actions">' +
        '<button class="copy-btn" onclick="copyMessageContent(\'' + msgId + '\')" title="Copy to clipboard">📋 Copy</button>' +
      '</div>' +
    '</div>';
  
  chatContainer.appendChild(messageDiv);
  scrollToBottom();
}

// Global function for copy button click
window.copyMessageContent = function(msgId) {
  var msgDiv = document.getElementById(msgId);
  if (msgDiv) {
    var rawText = msgDiv.getAttribute('data-raw-text');
    var copyBtn = msgDiv.querySelector('.copy-btn');
    if (rawText && copyBtn) {
      copyToClipboard(rawText, copyBtn);
    }
  }
};

function showLoading() {
  initializeElements();
  
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "message assistant";
  loadingDiv.id = "loading-message";
  loadingDiv.innerHTML = `
    <div class="message-avatar">AI</div>
    <div class="message-content">
      <div class="loading">
        <div class="loading-dot"></div>
        <div class="loading-dot"></div>
        <div class="loading-dot"></div>
      </div>
    </div>
  `;
  chatContainer.appendChild(loadingDiv);
  scrollToBottom();
}

function hideLoading() {
  const loadingMessage = document.getElementById("loading-message");
  if (loadingMessage) {
    loadingMessage.remove();
  }
}

function scrollToBottom() {
  if (chatContainer) {
    chatContainer.scrollTop = chatContainer.scrollHeight;
  }
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

/**
 * Simple markdown parser - converts markdown to HTML
 * Handles: bold, italic, headings, lists, code blocks, line breaks
 */
function parseMarkdown(text) {
  if (!text) return '';
  
  // First escape HTML to prevent XSS, but preserve structure
  var lines = text.split('\n');
  var result = [];
  var inCodeBlock = false;
  var inList = false;
  var listType = null;
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    
    // Code blocks
    if (line.trim().indexOf('```') === 0) {
      if (inCodeBlock) {
        result.push('</code></pre>');
        inCodeBlock = false;
      } else {
        if (inList) {
          result.push(listType === 'ol' ? '</ol>' : '</ul>');
          inList = false;
        }
        result.push('<pre><code>');
        inCodeBlock = true;
      }
      continue;
    }
    
    if (inCodeBlock) {
      result.push(escapeHtml(line));
      continue;
    }
    
    // Process inline formatting
    var processed = escapeHtml(line);
    
    // Bold **text** or __text__
    processed = processed.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    processed = processed.replace(/__(.+?)__/g, '<strong>$1</strong>');
    
    // Italic *text* or _text_
    processed = processed.replace(/\*([^*]+)\*/g, '<em>$1</em>');
    processed = processed.replace(/_([^_]+)_/g, '<em>$1</em>');
    
    // Inline code `code`
    processed = processed.replace(/`([^`]+)`/g, '<code class="inline-code">$1</code>');
    
    // Headings
    if (processed.match(/^#{1,6}\s/)) {
      if (inList) {
        result.push(listType === 'ol' ? '</ol>' : '</ul>');
        inList = false;
      }
      var level = processed.match(/^(#+)/)[1].length;
      var headingText = processed.replace(/^#+\s*/, '');
      result.push('<h' + (level + 2) + ' class="md-heading">' + headingText + '</h' + (level + 2) + '>');
      continue;
    }
    
    // Unordered list items (• or - or *)
    var unorderedMatch = processed.match(/^[\s]*[-•*]\s+(.+)$/);
    if (unorderedMatch) {
      if (!inList || listType !== 'ul') {
        if (inList) result.push('</ol>');
        result.push('<ul class="md-list">');
        inList = true;
        listType = 'ul';
      }
      result.push('<li>' + unorderedMatch[1] + '</li>');
      continue;
    }
    
    // Ordered list items (1. 2. etc)
    var orderedMatch = processed.match(/^[\s]*(\d+)\.\s+(.+)$/);
    if (orderedMatch) {
      if (!inList || listType !== 'ol') {
        if (inList) result.push('</ul>');
        result.push('<ol class="md-list">');
        inList = true;
        listType = 'ol';
      }
      result.push('<li>' + orderedMatch[2] + '</li>');
      continue;
    }
    
    // Close list if line doesn't continue it
    if (inList && processed.trim() !== '') {
      result.push(listType === 'ol' ? '</ol>' : '</ul>');
      inList = false;
    }
    
    // Empty line
    if (processed.trim() === '') {
      if (inList) {
        result.push(listType === 'ol' ? '</ol>' : '</ul>');
        inList = false;
      }
      result.push('<br>');
      continue;
    }
    
    // Regular paragraph
    result.push('<p class="md-paragraph">' + processed + '</p>');
  }
  
  // Close any open tags
  if (inCodeBlock) result.push('</code></pre>');
  if (inList) result.push(listType === 'ol' ? '</ol>' : '</ul>');
  
  return result.join('\n');
}

/**
 * Generate unique ID for messages
 */
var messageIdCounter = 0;
function generateMessageId() {
  return 'msg-' + (++messageIdCounter) + '-' + Date.now();
}

/**
 * Copy text to clipboard
 */
function copyToClipboard(text, buttonElement) {
  // Try modern clipboard API first
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(function() {
      showCopySuccess(buttonElement);
    }).catch(function() {
      fallbackCopyToClipboard(text, buttonElement);
    });
  } else {
    fallbackCopyToClipboard(text, buttonElement);
  }
}

function fallbackCopyToClipboard(text, buttonElement) {
  var textArea = document.createElement('textarea');
  textArea.value = text;
  textArea.style.position = 'fixed';
  textArea.style.left = '-9999px';
  document.body.appendChild(textArea);
  textArea.select();
  try {
    document.execCommand('copy');
    showCopySuccess(buttonElement);
  } catch (err) {
    console.error('Failed to copy:', err);
  }
  document.body.removeChild(textArea);
}

function showCopySuccess(buttonElement) {
  var originalText = buttonElement.innerHTML;
  buttonElement.innerHTML = '✓ Copied!';
  buttonElement.classList.add('copied');
  setTimeout(function() {
    buttonElement.innerHTML = originalText;
    buttonElement.classList.remove('copied');
  }, 2000);
}

async function sendMessage() {
  initializeElements();

  const message = messageInput.value.trim();
  if (!message) return;

  // Check for special commands first (starting with /)
  if (message.startsWith('/')) {
    addUserMessage(message); // Show the user's command
    messageInput.value = "";
    const handled = await handleSpecialCommand(message);
    if (handled) return;
  }
  
  // Show user message first
  addUserMessage(message);
  messageInput.value = "";
  
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
  // No more hardcoded pattern matching!

  // Check if API key is configured for selected provider
  const hasGroqKey = await apiKeyManager.hasApiKey('groq');
  const hasGeminiKey = await apiKeyManager.hasApiKey('gemini');

  if (!hasGroqKey && !hasGeminiKey) {
    addSystemMessage("❌ No API key configured. Please click the settings (⚙️) icon and add at least one provider's API key to use chat features.");
    messageInput.disabled = false;
    sendButton.disabled = false;
    return;
  }

  // If selected provider doesn't have a key, switch to available one
  if (currentProvider === 'groq' && !hasGroqKey) {
    if (hasGeminiKey) {
      currentProvider = 'gemini';
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

  // User message already shown above, just add to history
  chatHistory.push({ role: 'user', content: message });

  // Show loading indicator
  showLoading();

  try {
    // Determine if we need document context
    const needsDocContext = shouldIncludeDocumentContext(message);
    
    let systemContext = buildSystemContext();
    
    // Get document context if needed
    if (needsDocContext) {
      try {
        const docContext = await documentService.getDocumentContext();
        
        if (!docContext.isEmpty) {
          const formattedContext = documentService.formatContextForAI(docContext);
          systemContext += `\n\n${formattedContext}`;
        } else {
          systemContext += "\n\nNote: The document is currently empty.";
        }
      } catch (error) {
        console.warn("Could not get document context:", error);
        systemContext += "\n\nNote: Unable to read document at this time.";
      }
    }

    // Prepare messages for API
    const messages = [
      { role: 'system', content: systemContext },
      ...chatHistory.slice(-10) // Keep last 10 messages for context
    ];

    // Send to appropriate AI service
    let response;
    if (currentProvider === 'gemini') {
      response = await geminiService.sendMessage(messages);
    } else {
      response = await groqService.sendMessage(messages);
    }
    
    hideLoading();
    
    // Check for and execute any action commands in the response
    var processedResponse = await parseAndExecuteActions(response);
    
    // Add AI response to chat (with action tags removed)
    addAssistantMessage(processedResponse);
    
    // Add to chat history
    chatHistory.push({ role: 'assistant', content: response });
    
  } catch (error) {
    hideLoading();
    addAssistantMessage(`❌ Error: ${error.message}`);
    console.error("Error:", error);
  } finally {
    // Re-enable input
    messageInput.disabled = false;
    sendButton.disabled = false;
    messageInput.focus();
  }
}

/**
 * Determine if document context should be included based on user message
 * @param {string} message - User's message
 * @returns {boolean} True if document context is needed
 */
function shouldIncludeDocumentContext(message) {
  // Almost always include document context - let AI decide relevance
  // Only skip for very basic greetings
  const skipPatterns = /^(hi|hello|hey|thanks|thank you|ok|okay|yes|no|bye)[\s!.?]*$/i;
  return !skipPatterns.test(message.trim());
}

/**
 * Build system context for AI
 * @returns {string} System context prompt
 */
function buildSystemContext() {
  return `You are an AI assistant that edits Microsoft Word documents using ACTION commands.

## SUPPORTED ACTIONS (use ONLY these - no others exist):

### 1. REPLACE - Replace/reformat entire document content
[ACTION: REPLACE]
---CONTENT START---
# Title
## Section
Paragraph text
- Bullet point
| Col1 | Col2 |
| Data | Data |
---CONTENT END---

### 2. FORMAT - Apply formatting to specific text
[ACTION: FORMAT target="first heading" bold=true]

### 3. INSERT - Add content at end of document
[ACTION: INSERT heading="Section" content="text here" newpage=false]

### 4. CREATE - Create a NEW blank document (opens in new window)
[ACTION: CREATE]
---CONTENT START---
# New Document Title
Content for the new document...
---CONTENT END---

## FORMATTING SYNTAX (for REPLACE and CREATE):
- # = Heading 1, ## = Heading 2, ### = Heading 3
- Lines starting with - = bullet points
- Lines with | = table rows (pipe-separated)

## RULES:
1. ONLY use the 4 actions above - DO NOT invent new actions
2. For summarizing: use REPLACE to update current doc, or CREATE for new doc
3. CREATE opens a new Word window with the content
4. Keep response brief after the action`;
}

/**
 * Handle special commands (like /analyze, /summarize, etc.)
 * @param {string} message - User message
 * @returns {Promise<boolean>} True if command was handled
 */
async function handleSpecialCommand(message) {
  const command = message.toLowerCase().trim();
  
  if (command === '/analyze' || command === '/stats') {
    showLoading();
    try {
      const analysis = await documentService.analyzeDocument();
      hideLoading();
      
      if (analysis.isEmpty) {
        addAssistantMessage("📄 The document is empty. Add some content first!");
      } else {
        let response = `📊 **Document Analysis**\n\n`;
        response += `**Statistics:**\n`;
        response += `- Words: ${analysis.statistics.wordCount}\n`;
        response += `- Characters: ${analysis.statistics.characterCount}\n`;
        response += `- Sentences: ${analysis.statistics.sentenceCount}\n`;
        response += `- Paragraphs: ${analysis.statistics.paragraphCount}\n`;
        response += `- Reading Time: ~${analysis.readingTime} minute(s)\n`;
        response += `- Avg Words/Sentence: ${analysis.averageWordsPerSentence}\n\n`;
        
        if (analysis.topWords && analysis.topWords.length > 0) {
          response += `**Most Common Words:**\n`;
          analysis.topWords.slice(0, 5).forEach((item, i) => {
            response += `${i + 1}. "${item.word}" (${item.count}x)\n`;
          });
        }
        
        addAssistantMessage(response);
      }
      return true;
    } catch (error) {
      hideLoading();
      addAssistantMessage(`❌ Error analyzing document: ${error.message}`);
      return true;
    }
  }
  
  if (command === '/help') {
    const helpText = `🤖 **AI Helper Commands**\n\n` +
      `**Slash Commands (select text first):**\n` +
      `• /bold, /italic, /underline\n` +
      `• /center, /left, /right\n` +
      `• /h1, /h2, /h3 - Heading styles\n` +
      `• /analyze - Document stats\n\n` +
      `**Natural Language (no selection needed!):**\n` +
      `• "make the first heading bold"\n` +
      `• "center the title"\n` +
      `• "make 'What is Android?' italic"\n` +
      `• "underline the heading"\n\n` +
      `**With Selection:**\n` +
      `• "make it bold and italic"\n` +
      `• "center this"\n\n` +
      `💡 Just describe what you want - I'll do it!`;
    addAssistantMessage(helpText);
    return true;
  }
  
  // Formatting commands - require text selection
  if (command === '/bold') {
    return await applyFormatting({ bold: true }, "bold");
  }
  if (command === '/italic') {
    return await applyFormatting({ italic: true }, "italic");
  }
  if (command === '/underline') {
    return await applyFormatting({ underline: true }, "underlined");
  }
  if (command === '/center') {
    return await applyAlignment("Center", "centered");
  }
  if (command === '/left') {
    return await applyAlignment("Left", "left-aligned");
  }
  if (command === '/right') {
    return await applyAlignment("Right", "right-aligned");
  }
  if (command === '/h1') {
    return await applyHeadingStyle(1);
  }
  if (command === '/h2') {
    return await applyHeadingStyle(2);
  }
  if (command === '/h3') {
    return await applyHeadingStyle(3);
  }
  
  return false;
}

/**
 * Apply formatting to selected text
 */
async function applyFormatting(options, description) {
  try {
    await documentService.formatSelection(options);
    addAssistantMessage("✅ Done! Made the selected text " + description + ".");
    return true;
  } catch (error) {
    if (error.message.includes("No text selected")) {
      addAssistantMessage("⚠️ Please select some text in your document first, then try again.");
    } else {
      addAssistantMessage("❌ Error: " + error.message);
    }
    return true;
  }
}

/**
 * Apply alignment to selected paragraphs
 */
async function applyAlignment(alignment, description) {
  try {
    await documentService.setAlignment(alignment);
    addAssistantMessage("✅ Done! Text is now " + description + ".");
    return true;
  } catch (error) {
    addAssistantMessage("❌ Error: " + error.message);
    return true;
  }
}

/**
 * Apply heading style to selected text
 */
async function applyHeadingStyle(level) {
  try {
    await documentService.applyHeading(level);
    addAssistantMessage("✅ Done! Applied Heading " + level + " style.");
    return true;
  } catch (error) {
    addAssistantMessage("❌ Error: " + error.message);
    return true;
  }
}

/**
 * Parse AI response for action commands and execute them
 * @param {string} response - AI response text
 * @returns {string} Response with action tags removed
 */
async function parseAndExecuteActions(response) {
  var cleanedResponse = response;
  
  // Handle FORMAT action - support multiple FORMAT actions in one response
  var formatRegex = /\[ACTION:\s*FORMAT\s+([^\]]+)\]/gi;
  var formatMatch;
  
  while ((formatMatch = formatRegex.exec(response)) !== null) {
    var actionParams = formatMatch[1];
    
    // Parse parameters
    var targetMatch = actionParams.match(/target\s*=\s*["']([^"']+)["']/i);
    var target = targetMatch ? targetMatch[1] : null;
    
    // Parse formatting options - check for both true AND false values
    var formatOptions = {};
    
    // Bold
    if (/bold\s*=\s*true/i.test(actionParams)) {
      formatOptions.bold = true;
    } else if (/bold\s*=\s*false/i.test(actionParams)) {
      formatOptions.bold = false;
    }
    
    // Italic
    if (/italic\s*=\s*true/i.test(actionParams)) {
      formatOptions.italic = true;
    } else if (/italic\s*=\s*false/i.test(actionParams)) {
      formatOptions.italic = false;
    }
    
    // Underline
    if (/underline\s*=\s*true/i.test(actionParams)) {
      formatOptions.underline = true;
    } else if (/underline\s*=\s*false/i.test(actionParams)) {
      formatOptions.underline = false;
    }
    
    // Alignment
    var wantsCenter = /center\s*=\s*true/i.test(actionParams);
    var wantsLeft = /left\s*=\s*true/i.test(actionParams);
    var wantsRight = /right\s*=\s*true/i.test(actionParams);
    
    if (target) {
      try {
        var isHeadingTarget = /first heading|the heading|title|the title/i.test(target);
        
        if (isHeadingTarget) {
          // Format first heading
          if (Object.keys(formatOptions).length > 0) {
            await documentService.formatFirstHeading(formatOptions);
          }
          
          if (wantsCenter) {
            await documentService.alignFirstHeading("Center");
          } else if (wantsLeft) {
            await documentService.alignFirstHeading("Left");
          } else if (wantsRight) {
            await documentService.alignFirstHeading("Right");
          }
        } else {
          // Format specific text
          if (Object.keys(formatOptions).length > 0) {
            await documentService.formatText(target, formatOptions);
          }
        }
        
        console.log("FORMAT action executed successfully for target:", target, "options:", formatOptions);
      } catch (error) {
        console.error("Error executing FORMAT action:", error);
      }
    }
  }
  
  // Remove all FORMAT actions from displayed response
  cleanedResponse = cleanedResponse.replace(/\[ACTION:\s*FORMAT\s+[^\]]+\]/gi, '').trim();
  
  // Handle INSERT action
  var insertRegex = /\[ACTION:\s*INSERT\s+([^\]]+)\]/gi;
  var insertMatch = insertRegex.exec(response);
  
  if (insertMatch) {
    var insertParams = insertMatch[1];
    
    // Parse parameters - handle multiline content
    var headingMatch = insertParams.match(/heading\s*=\s*["']([^"']+)["']/i);
    var heading = headingMatch ? headingMatch[1] : null;
    
    // Content can be multiline, so we need special handling
    var contentMatch = insertParams.match(/content\s*=\s*["'](.+?)["']\s*(?:newpage|$)/is);
    var content = contentMatch ? contentMatch[1] : null;
    
    var newPage = /newpage\s*=\s*true/i.test(insertParams);
    // Default to true if not specified
    if (!/newpage\s*=/i.test(insertParams)) {
      newPage = true;
    }
    
    if (content) {
      try {
        // Clean up content - replace escaped newlines with actual newlines
        content = content.replace(/\\n/g, '\n').trim();
        
        await documentService.insertContentSection(heading, content, newPage);
        console.log("INSERT action executed successfully");
      } catch (error) {
        console.error("Error executing INSERT action:", error);
      }
    } else {
      console.warn("INSERT action found but no content specified");
    }
    
    cleanedResponse = cleanedResponse.replace(insertRegex, '').trim();
  }
  
  // Handle TABLE action
  var tableRegex = /\[ACTION:\s*TABLE(?:\s+title\s*=\s*["']([^"']+)["'])?\s*\]\s*([\s\S]*?)\s*\[\/TABLE\]/gi;
  var tableMatch;
  
  while ((tableMatch = tableRegex.exec(response)) !== null) {
    var tableTitle = tableMatch[1] || null;
    var tableContent = tableMatch[2].trim();
    
    if (tableContent) {
      try {
        // Parse the table content - rows separated by newlines, columns by |
        var lines = tableContent.split('\n').filter(function(line) {
          return line.trim().length > 0;
        });
        
        if (lines.length > 0) {
          // First line is headers
          var headers = lines[0].split('|').map(function(h) {
            return h.trim();
          }).filter(function(h) {
            return h.length > 0;
          });
          
          // Remaining lines are data rows
          var rows = [];
          for (var i = 1; i < lines.length; i++) {
            var row = lines[i].split('|').map(function(cell) {
              return cell.trim();
            }).filter(function(cell) {
              return cell.length > 0 || rows.length === 0; // Allow empty cells
            });
            if (row.length > 0) {
              rows.push(row);
            }
          }
          
          if (headers.length > 0 && rows.length > 0) {
            await documentService.insertTable(headers, rows, tableTitle);
            console.log("TABLE action executed successfully:", headers.length, "cols,", rows.length, "rows");
            addSystemMessage("📊 Table inserted successfully!");
          }
        }
      } catch (error) {
        console.error("Error executing TABLE action:", error);
        addSystemMessage("⚠️ Failed to insert table: " + error.message);
      }
    }
  }
  
  // Remove all TABLE actions from displayed response
  cleanedResponse = cleanedResponse.replace(/\[ACTION:\s*TABLE(?:\s+[^\]]+)?\s*\][\s\S]*?\[\/TABLE\]/gi, '').trim();
  
  // Handle CREATE action (create new document with content)
  var createRegex = /\[ACTION:\s*CREATE\s*\]\s*---CONTENT START---\s*([\s\S]*?)\s*---CONTENT END---/gi;
  var createMatch = createRegex.exec(response);
  
  // Try lenient format for CREATE too
  if (!createMatch) {
    var lenientCreateRegex = /\[ACTION:\s*CREATE\s*\]\s*---CONTENT START---\s*([\s\S]+)/gi;
    createMatch = lenientCreateRegex.exec(response);
    if (createMatch) {
      var content = createMatch[1];
      content = content.replace(/\n\n\*\*Note:?\*\*[\s\S]*$/i, '');
      content = content.replace(/\n\nI've (created|made)[\s\S]*$/i, '');
      createMatch[1] = content.trim();
    }
  }
  
  if (createMatch) {
    var newDocContent = createMatch[1].trim();
    
    // Clean up any stray ACTION markers from the content
    newDocContent = newDocContent.replace(/\[ACTION:[^\]]*\]/gi, '');
    newDocContent = newDocContent.replace(/\[\/[A-Z]+\]/gi, '');
    newDocContent = newDocContent.trim();
    
    if (newDocContent && newDocContent.length > 10) {
      try {
        var result = await documentService.createDocument(newDocContent);
        if (result.success) {
          console.log("CREATE action executed successfully");
          addSystemMessage("📄 New document created! Check the new Word window.");
          
          // If content was provided, let user know they may need to paste it
          if (result.hasContent) {
            addSystemMessage("💡 Tip: The content for your new document is ready. You may need to paste it in the new window.");
          }
        }
      } catch (error) {
        console.error("Error executing CREATE action:", error);
        if (error.message === "CREATE_NOT_SUPPORTED") {
          // Fallback: offer to replace current document instead
          addSystemMessage("⚠️ Creating new documents isn't supported in this version of Word. Would you like me to REPLACE the current document content instead? Reply 'yes' to confirm.");
          // Store the content for potential REPLACE
          window._pendingCreateContent = newDocContent;
        } else {
          addSystemMessage("⚠️ Failed to create new document: " + error.message);
        }
      }
    }
    
    // Remove the CREATE action from displayed response
    cleanedResponse = cleanedResponse.replace(/\[ACTION:\s*CREATE\s*\]\s*---CONTENT START---[\s\S]*/gi, '').trim();
  }
  
  // Handle REPLACE action (replace entire document content)
  // Try strict format first (with END tag)
  var replaceRegex = /\[ACTION:\s*REPLACE\s*\]\s*---CONTENT START---\s*([\s\S]*?)\s*---CONTENT END---/gi;
  var replaceMatch = replaceRegex.exec(response);
  
  // If strict format not found, try lenient format (without END tag - content goes to end of response)
  if (!replaceMatch) {
    var lenientRegex = /\[ACTION:\s*REPLACE\s*\]\s*---CONTENT START---\s*([\s\S]+)/gi;
    replaceMatch = lenientRegex.exec(response);
    if (replaceMatch) {
      // Clean up: remove any trailing notes or messages after the content
      var content = replaceMatch[1];
      // Remove common ending patterns like "Note:", "I've reformatted", etc.
      content = content.replace(/\n\n\*\*Note:?\*\*[\s\S]*$/i, '');
      content = content.replace(/\n\nI've (reformatted|corrected|fixed)[\s\S]*$/i, '');
      content = content.replace(/\n\nPlease (let me know|confirm)[\s\S]*$/i, '');
      replaceMatch[1] = content.trim();
    }
  }
  
  if (replaceMatch) {
    var newContent = replaceMatch[1].trim();
    
    // Clean up any stray ACTION markers from the content
    newContent = newContent.replace(/\[ACTION:\s*TABLE[^\]]*\]/gi, '');
    newContent = newContent.replace(/\[\/TABLE\]/gi, '');
    newContent = newContent.replace(/\[ACTION:\s*FORMAT[^\]]*\]/gi, '');
    newContent = newContent.replace(/\[\/ACTION\]/gi, '');
    newContent = newContent.trim();
    
    if (newContent && newContent.length > 50) {
      // Show confirmation dialog before replacing
      var confirmed = await showReplaceConfirmation(newContent);
      
      if (confirmed) {
        try {
          await documentService.replaceDocumentContent(newContent);
          console.log("REPLACE action executed successfully, content length:", newContent.length);
          addSystemMessage("📝 Document has been reformatted!");
        } catch (error) {
          console.error("Error executing REPLACE action:", error);
          addSystemMessage("⚠️ Failed to apply changes to document.");
        }
      } else {
        addSystemMessage("❌ Document replacement cancelled.");
      }
    } else {
      console.warn("REPLACE action found but content was empty or too short:", newContent ? newContent.length : 0);
    }
    
    // Remove the action from displayed response
    cleanedResponse = cleanedResponse.replace(/\[ACTION:\s*REPLACE\s*\]\s*---CONTENT START---[\s\S]*/gi, '').trim();
  }
  
  // Final cleanup: remove ANY remaining action-like patterns that users shouldn't see
  cleanedResponse = cleanedResponse
    .replace(/\[ACTION:[^\]]*\]/gi, '')           // [ACTION: anything]
    .replace(/\[\/[A-Z]+\]/gi, '')                 // [/TABLE], [/ACTION], etc.
    .replace(/\[TOC\]/gi, '')                      // [TOC]
    .replace(/---CONTENT START---/gi, '')
    .replace(/---CONTENT END---/gi, '')
    .replace(/^\s*[-=]{3,}\s*$/gm, '')             // Lines with just --- or ===
    .replace(/\n{3,}/g, '\n\n')                    // Multiple newlines to double
    .trim();
  
  return cleanedResponse;
}

/**
 * Show confirmation dialog before replacing document content
 * @param {string} newContent - The new content to be applied
 * @returns {Promise<boolean>} True if user confirms, false otherwise
 */
function showReplaceConfirmation(newContent) {
  return new Promise(function(resolve) {
    // Create modal overlay
    var overlay = document.createElement('div');
    overlay.id = 'replace-confirm-overlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:10000;display:flex;align-items:center;justify-content:center;';
    
    // Create modal content
    var modal = document.createElement('div');
    modal.style.cssText = 'background:white;border-radius:8px;max-width:90%;max-height:80%;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 4px 20px rgba(0,0,0,0.3);';
    
    // Header
    var header = document.createElement('div');
    header.style.cssText = 'padding:16px;background:#f0f0f0;border-bottom:1px solid #ddd;';
    header.innerHTML = '<h3 style="margin:0;font-size:16px;">⚠️ Review Document Changes</h3><p style="margin:8px 0 0;font-size:12px;color:#666;">The AI wants to replace your document with the following content. Please review before applying.</p>';
    
    // Preview area
    var preview = document.createElement('div');
    preview.style.cssText = 'padding:16px;overflow-y:auto;max-height:300px;font-family:monospace;font-size:11px;white-space:pre-wrap;background:#fafafa;border-bottom:1px solid #ddd;';
    
    // Truncate preview if too long
    var previewText = newContent;
    if (previewText.length > 3000) {
      previewText = previewText.substring(0, 3000) + '\n\n... [' + (newContent.length - 3000) + ' more characters]';
    }
    preview.textContent = previewText;
    
    // Stats
    var stats = document.createElement('div');
    stats.style.cssText = 'padding:8px 16px;background:#e8f4fd;font-size:11px;color:#0066cc;';
    var lineCount = newContent.split('\n').length;
    var wordCount = newContent.trim().split(/\s+/).length;
    stats.textContent = '📊 ' + wordCount + ' words, ' + lineCount + ' lines, ' + newContent.length + ' characters';
    
    // Buttons
    var buttons = document.createElement('div');
    buttons.style.cssText = 'padding:16px;display:flex;gap:12px;justify-content:flex-end;';
    
    var cancelBtn = document.createElement('button');
    cancelBtn.textContent = '❌ Cancel';
    cancelBtn.style.cssText = 'padding:10px 20px;border:1px solid #ccc;background:#fff;border-radius:4px;cursor:pointer;font-size:14px;';
    cancelBtn.onclick = function() {
      document.body.removeChild(overlay);
      resolve(false);
    };
    
    var applyBtn = document.createElement('button');
    applyBtn.textContent = '✅ Apply Changes';
    applyBtn.style.cssText = 'padding:10px 20px;border:none;background:#0078d4;color:white;border-radius:4px;cursor:pointer;font-size:14px;';
    applyBtn.onclick = function() {
      document.body.removeChild(overlay);
      resolve(true);
    };
    
    buttons.appendChild(cancelBtn);
    buttons.appendChild(applyBtn);
    
    modal.appendChild(header);
    modal.appendChild(preview);
    modal.appendChild(stats);
    modal.appendChild(buttons);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    
    // Focus the cancel button by default (safer)
    cancelBtn.focus();
  });
}

// Office.js document operations (legacy - using DocumentService now)
async function readDocumentContent() {
  return documentService.readDocumentText();
}

async function insertText(text) {
  return documentService.insertText(text, "End");
}

// API Key Setup Functions
function showApiKeySetup() {
  removeWelcomeScreen();
  initializeElements();
  
  isSettingsOpen = true;
  
  const setupDiv = document.createElement("div");
  setupDiv.id = "api-key-setup";
  setupDiv.className = "api-key-setup";
  setupDiv.innerHTML = `
    <div class="setup-content">
      <h2>🔑 Welcome to AI Helper!</h2>
      <p>Choose an AI provider and get your API key to get started.</p>
      
      <div class="provider-tabs">
        <button class="provider-tab active" onclick="switchSetupProvider('groq')">
          ⚡ Groq (Llama 3.1)
        </button>
        <button class="provider-tab" onclick="switchSetupProvider('gemini')">
          🧠 Google Gemini
        </button>
      </div>

      <div id="groq-setup" class="provider-setup active">
        <div class="setup-steps">
          <div class="setup-step">
            <strong>Step 1:</strong> Create a free account at Groq
            <br>
            <button onclick="window.open('https://console.groq.com', '_blank')" class="link-button">
              Open Groq Console →
            </button>
          </div>
          
          <div class="setup-step">
            <strong>Step 2:</strong> Navigate to API Keys section
          </div>
          
          <div class="setup-step">
            <strong>Step 3:</strong> Create a new API key and copy it
          </div>
          
          <div class="setup-step">
            <strong>Step 4:</strong> Paste your API key below
            <input 
              type="password" 
              id="groq-api-key-input" 
              placeholder="gsk_..." 
              class="api-key-input"
            />
            <button id="groq-show-key-btn" class="show-key-btn" onclick="toggleApiKeyVisibility('groq')">
              👁️ Show
            </button>
          </div>
        </div>
        
        <div id="groq-api-key-error" class="api-key-error"></div>
        
        <div class="setup-buttons">
          <button id="groq-test-key-btn" class="secondary-button">Test Connection</button>
          <button id="groq-save-key-btn" class="primary-button">Save Groq Key</button>
        </div>
      </div>

      <div id="gemini-setup" class="provider-setup">
        <div class="setup-steps">
          <div class="setup-step">
            <strong>Step 1:</strong> Go to Google AI Studio
            <br>
            <button onclick="window.open('https://aistudio.google.com/apikey', '_blank')" class="link-button">
              Open Google AI Studio →
            </button>
          </div>
          
          <div class="setup-step">
            <strong>Step 2:</strong> Create a new API key (or use existing)
          </div>
          
          <div class="setup-step">
            <strong>Step 3:</strong> Copy your API key from Google
          </div>
          
          <div class="setup-step">
            <strong>Step 4:</strong> Paste your API key below
            <input 
              type="password" 
              id="gemini-api-key-input" 
              placeholder="AIza..." 
              class="api-key-input"
            />
            <button id="gemini-show-key-btn" class="show-key-btn" onclick="toggleApiKeyVisibility('gemini')">
              👁️ Show
            </button>
          </div>
        </div>
        
        <div id="gemini-api-key-error" class="api-key-error"></div>
        
        <div class="setup-buttons">
          <button id="gemini-test-key-btn" class="secondary-button">Test Connection</button>
          <button id="gemini-save-key-btn" class="primary-button">Save Gemini Key</button>
        </div>
      </div>
      
      <p class="privacy-note">
        🔒 Your API keys are stored securely and never shared.
      </p>
    </div>
  `;
  
  chatContainer.appendChild(setupDiv);
  
  // Add event listeners for Groq
  document.getElementById("groq-test-key-btn").onclick = () => testApiKey('groq');
  document.getElementById("groq-save-key-btn").onclick = () => saveApiKey('groq');
  
  // Add event listeners for Gemini
  document.getElementById("gemini-test-key-btn").onclick = () => testApiKey('gemini');
  document.getElementById("gemini-save-key-btn").onclick = () => saveApiKey('gemini');
  
  document.getElementById("groq-api-key-input").focus();
}

function switchSetupProvider(provider) {
  // Hide all provider setups
  document.getElementById('groq-setup').classList.remove('active');
  document.getElementById('gemini-setup').classList.remove('active');
  
  // Remove active class from all tabs
  document.querySelectorAll('.provider-tab').forEach(tab => tab.classList.remove('active'));
  
  // Show selected provider
  document.getElementById(`${provider}-setup`).classList.add('active');
  
  // Mark tab as active
  event.target.classList.add('active');
  
  // Focus on input
  const input = document.getElementById(`${provider}-api-key-input`);
  if (input) input.focus();
}

function toggleApiKeyVisibility(provider) {
  const input = document.getElementById(`${provider}-api-key-input`);
  const btn = document.getElementById(`${provider}-show-key-btn`);
  if (input.type === "password") {
    input.type = "text";
    btn.textContent = "🙈 Hide";
  } else {
    input.type = "password";
    btn.textContent = "👁️ Show";
  }
}

async function testApiKey(provider) {
  const input = document.getElementById(`${provider}-api-key-input`);
  const errorDiv = document.getElementById(`${provider}-api-key-error`);
  const testBtn = document.getElementById(`${provider}-test-key-btn`);
  const apiKey = input.value.trim();
  
  errorDiv.textContent = "";
  errorDiv.className = "api-key-error";
  
  if (!apiKey) {
    errorDiv.textContent = "⚠️ Please enter an API key";
    errorDiv.className = "api-key-error error";
    return;
  }
  
  // Validate format
  if (!apiKeyManager.validateFormat(apiKey, provider)) {
    const format = provider === 'groq' ? "gsk_" : "AIza";
    errorDiv.textContent = `⚠️ Invalid API key format. ${provider === 'groq' ? 'Groq' : 'Google'} keys start with '${format}'`;
    errorDiv.className = "api-key-error error";
    return;
  }
  
  testBtn.disabled = true;
  testBtn.textContent = "Testing...";
  
  try {
    let result;
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
      errorDiv.textContent = `❌ ${result.error}`;
      errorDiv.className = "api-key-error error";
    }
  } catch (error) {
    errorDiv.textContent = `❌ Error: ${error.message}`;
    errorDiv.className = "api-key-error error";
  } finally {
    testBtn.disabled = false;
    testBtn.textContent = "Test Connection";
  }
}

async function saveApiKey(provider) {
  const input = document.getElementById(`${provider}-api-key-input`);
  const errorDiv = document.getElementById(`${provider}-api-key-error`);
  const saveBtn = document.getElementById(`${provider}-save-key-btn`);
  const apiKey = input.value.trim();
  
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
    // Save the API key
    let saved;
    if (provider === 'groq') {
      saved = await apiKeyManager.saveGroqApiKey(apiKey);
      groqService.setApiKey(apiKey);
    } else {
      saved = await apiKeyManager.saveGeminiApiKey(apiKey);
      geminiService.setApiKey(apiKey);
    }
    
    if (saved) {
      // Set as active provider if first key
      const otherProvider = provider === 'groq' ? 'gemini' : 'groq';
      const hasOtherKey = provider === 'groq' 
        ? await apiKeyManager.hasApiKey('gemini')
        : await apiKeyManager.hasApiKey('groq');
      
      if (!hasOtherKey) {
        currentProvider = provider;
        await apiKeyManager.setActiveProvider(provider);
      }
      
      // Remove setup UI
      const setupDiv = document.getElementById("api-key-setup");
      if (setupDiv) {
        setupDiv.remove();
      }
      
      // Show success message
      addSystemMessage(`✅ ${apiKeyManager.getProviderName(provider)} key saved successfully!`);
      addAssistantMessage(`Great! I'm ready to chat using ${apiKeyManager.getProviderName(provider)}.\n\nI can help you with:\n\n• Summarizing documents\n• Editing and formatting text\n• Answering questions about your document\n• Creating tables, headers, and more\n\nWhat would you like to do?`);
    } else {
      errorDiv.textContent = `❌ Failed to save ${provider === 'groq' ? 'Groq' : 'Gemini'} API key`;
      errorDiv.className = "api-key-error error";
    }
  } catch (error) {
    errorDiv.textContent = `❌ Error: ${error.message}`;
    errorDiv.className = "api-key-error error";
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = `Save ${provider === 'groq' ? 'Groq' : 'Gemini'} Key`;
  }
}

// Provider switch functions removed - now handled in settings panel

// Track if settings panel is open
let isSettingsOpen = false;

function toggleSettings() {
  if (isSettingsOpen) {
    closeSettingsPanel();
  } else {
    showApiKeySettings();
  }
}

async function showApiKeySettings() {
  // Check if we already have keys configured
  const hasGroqKey = await apiKeyManager.hasApiKey('groq');
  const hasGeminiKey = await apiKeyManager.hasApiKey('gemini');
  
  if (hasGroqKey || hasGeminiKey) {
    // Show settings panel with current configuration
    showSettingsPanel(hasGroqKey, hasGeminiKey);
  } else {
    // No keys - show initial setup
    showApiKeySetup();
  }
}

async function showSettingsPanel(hasGroqKey, hasGeminiKey) {
  removeWelcomeScreen();
  initializeElements();
  
  // Remove any existing settings panel
  const existingPanel = document.getElementById('settings-panel');
  if (existingPanel) {
    existingPanel.remove();
  }
  
  isSettingsOpen = true;
  
  const currentProviderName = apiKeyManager.getProviderName(currentProvider);
  
  const panelDiv = document.createElement('div');
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
  document.getElementById('select-groq-btn').onclick = () => switchActiveProvider('groq');
  document.getElementById('select-gemini-btn').onclick = () => switchActiveProvider('gemini');
  document.getElementById('save-groq-key-btn').onclick = () => saveKeyFromSettings('groq');
  document.getElementById('save-gemini-key-btn').onclick = () => saveKeyFromSettings('gemini');
  document.getElementById('close-settings-btn').onclick = closeSettingsPanel;
}

async function switchActiveProvider(provider) {
  const hasKey = await apiKeyManager.hasApiKey(provider);
  if (!hasKey) {
    const statusDiv = document.getElementById('settings-status');
    statusDiv.textContent = `⚠️ Please add a ${apiKeyManager.getProviderName(provider)} API key first`;
    statusDiv.className = 'api-key-error error';
    return;
  }
  
  currentProvider = provider;
  await apiKeyManager.setActiveProvider(provider);
  
  // Update UI
  const groqBtn = document.getElementById('select-groq-btn');
  const geminiBtn = document.getElementById('select-gemini-btn');
  
  if (provider === 'groq') {
    groqBtn.className = 'primary-button';
    geminiBtn.className = 'secondary-button';
  } else {
    groqBtn.className = 'secondary-button';
    geminiBtn.className = 'primary-button';
  }
  
  const statusDiv = document.getElementById('settings-status');
  statusDiv.textContent = `✅ Switched to ${apiKeyManager.getProviderName(provider)}`;
  statusDiv.className = 'api-key-error success';
}

async function saveKeyFromSettings(provider) {
  const inputId = provider === 'groq' ? 'settings-groq-key' : 'settings-gemini-key';
  const input = document.getElementById(inputId);
  const statusDiv = document.getElementById('settings-status');
  const apiKey = input.value.trim();
  
  if (!apiKey) {
    statusDiv.textContent = '⚠️ Please enter an API key';
    statusDiv.className = 'api-key-error error';
    return;
  }
  
  if (!apiKeyManager.validateFormat(apiKey, provider)) {
    const format = provider === 'groq' ? 'gsk_' : 'AIza';
    statusDiv.textContent = `⚠️ Invalid format. ${apiKeyManager.getProviderName(provider)} keys start with '${format}'`;
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
    
    statusDiv.textContent = `✅ ${apiKeyManager.getProviderName(provider)} key saved!`;
    statusDiv.className = 'api-key-error success';
    input.value = '';
    input.placeholder = '••••••••••••••••';
    
    // Enable the provider button
    const btnId = provider === 'groq' ? 'select-groq-btn' : 'select-gemini-btn';
    const btn = document.getElementById(btnId);
    btn.disabled = false;
    btn.style.opacity = '1';
    btn.innerHTML = provider === 'groq' ? '⚡ Groq ✓' : '🧠 Gemini ✓';
    
  } catch (error) {
    statusDiv.textContent = `❌ Error: ${error.message}`;
    statusDiv.className = 'api-key-error error';
  }
}

function closeSettingsPanel() {
  const panel = document.getElementById('settings-panel');
  if (panel) {
    panel.remove();
  }
  // Also close api-key-setup if open
  const setupPanel = document.getElementById('api-key-setup');
  if (setupPanel) {
    setupPanel.remove();
  }
  isSettingsOpen = false;
}

async function getDocumentContext() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error.message);
      }
    });
  });
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
  const debugEl = document.getElementById('debug-output');
  if (debugEl) {
    debugEl.textContent = 'ERROR: ' + bundleError.message;
    debugEl.style.color = 'red';
    debugEl.style.background = '#ffe0e0';
  }
}
