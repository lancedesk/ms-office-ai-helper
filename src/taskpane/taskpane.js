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
  
  // Check for natural language formatting commands (e.g., "make it bold")
  addUserMessage(message); // Always show user message first
  messageInput.value = "";
  
  const formattingHandled = await handleNaturalFormattingCommand(message);
  if (formattingHandled) {
    return;
  }

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
  const keywords = [
    'document', 'doc', 'text', 'content', 'write', 'written',
    'read', 'reading', 'summarize', 'summary', 'analyze', 'analysis',
    'this', 'it', 'what', 'about', 'tell me', 'explain',
    'format', 'edit', 'change', 'modify', 'improve',
    'paragraph', 'sentence', 'word', 'section', 'chapter',
    'heading', 'header', 'title', 'structure'
  ];
  
  const lowerMessage = message.toLowerCase();
  return keywords.some(keyword => lowerMessage.includes(keyword));
}

/**
 * Build system context for AI
 * @returns {string} System context prompt
 */
function buildSystemContext() {
  return `You are an intelligent AI assistant integrated into Microsoft Word with the ability to DIRECTLY EDIT the document.

## Your Capabilities:
1. **Read & Analyze**: Read document content, provide summaries, answer questions
2. **Format Text**: Apply bold, italic, underline, alignment, headings
3. **Rewrite/Reformat**: Completely rewrite or reformat the document with proper structure
4. **Insert Content**: Add new sections to the document

## CRITICAL - You MUST Use Action Commands:
When the user asks you to FORMAT, EDIT, REFORMAT, FIX, or MODIFY the document, you MUST include an ACTION command. DO NOT just show the corrected text - use an action to apply it!

### ACTION: FORMAT (style specific text)
[ACTION: FORMAT target="text to find" bold=true italic=true underline=true center=true]

Example: User says "make the first heading bold"
→ You respond: "Done! I've made the heading bold. [ACTION: FORMAT target="first heading" bold=true]"

### ACTION: REPLACE (replace entire document with reformatted content)
Use this when user asks to "fix formatting", "reformat the document", "correct the document", etc.
[ACTION: REPLACE]
---CONTENT START---
Your reformatted document content here...
Use proper structure with headings, paragraphs, lists.
---CONTENT END---

Example: User says "fix the formatting issues in this document"
→ You respond: "I've reformatted the document with proper headings, spacing, and structure.
[ACTION: REPLACE]
---CONTENT START---
What is Android?

Android is an open-source operating system...

Features of Android

• Beautiful UI - intuitive user interface
• Connectivity - supports WiFi, Bluetooth, NFC
...
---CONTENT END---"

### ACTION: INSERT (add new content at end)
[ACTION: INSERT heading="Section Title" content="Content here..." newpage=true]

Example: User says "add a summary to the document"
→ You respond: "I've added a summary section. [ACTION: INSERT heading="Summary" content="This document covers..." newpage=true]"

## IMPORTANT RULES:
1. ALWAYS use an ACTION command when the user wants document changes - never just show text!
2. For "fix formatting" or "reformat" requests, use REPLACE to rewrite the whole document properly
3. For small changes to specific text, use FORMAT
4. For adding new content, use INSERT
5. If you're not sure what to do, ask the user
6. Be concise in your confirmation message`;
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
  
  // Handle FORMAT action
  var formatRegex = /\[ACTION:\s*FORMAT\s+([^\]]+)\]/gi;
  var formatMatch = formatRegex.exec(response);
  
  if (formatMatch) {
    var actionParams = formatMatch[1];
    
    // Parse parameters
    var targetMatch = actionParams.match(/target\s*=\s*["']([^"']+)["']/i);
    var target = targetMatch ? targetMatch[1] : null;
    
    var wantsBold = /bold\s*=\s*true/i.test(actionParams);
    var wantsItalic = /italic\s*=\s*true/i.test(actionParams);
    var wantsUnderline = /underline\s*=\s*true/i.test(actionParams);
    var wantsCenter = /center\s*=\s*true/i.test(actionParams);
    var wantsLeft = /left\s*=\s*true/i.test(actionParams);
    var wantsRight = /right\s*=\s*true/i.test(actionParams);
    
    if (target) {
      try {
        var isHeadingTarget = /first heading|the heading|title|the title/i.test(target);
        
        if (isHeadingTarget) {
          // Format first heading
          var formatOptions = {};
          if (wantsBold) formatOptions.bold = true;
          if (wantsItalic) formatOptions.italic = true;
          if (wantsUnderline) formatOptions.underline = true;
          
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
          var formatOptions = {};
          if (wantsBold) formatOptions.bold = true;
          if (wantsItalic) formatOptions.italic = true;
          if (wantsUnderline) formatOptions.underline = true;
          
          if (Object.keys(formatOptions).length > 0) {
            await documentService.formatText(target, formatOptions);
          }
        }
        
        console.log("FORMAT action executed successfully for target:", target);
      } catch (error) {
        console.error("Error executing FORMAT action:", error);
      }
    }
    
    cleanedResponse = cleanedResponse.replace(formatRegex, '').trim();
  }
  
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
    
    if (newContent && newContent.length > 50) {
      // Only replace if content is substantial (more than 50 chars)
      try {
        await documentService.replaceDocumentContent(newContent);
        console.log("REPLACE action executed successfully, content length:", newContent.length);
        addSystemMessage("📝 Document has been reformatted!");
      } catch (error) {
        console.error("Error executing REPLACE action:", error);
        addSystemMessage("⚠️ Failed to apply changes to document.");
      }
    } else {
      console.warn("REPLACE action found but content was empty or too short:", newContent ? newContent.length : 0);
    }
    
    // Remove the action from displayed response
    cleanedResponse = cleanedResponse.replace(/\[ACTION:\s*REPLACE\s*\]\s*---CONTENT START---[\s\S]*/gi, '').trim();
  }
  
  return cleanedResponse;
}

/**
 * Check if message contains formatting intent and execute it
 */
async function handleNaturalFormattingCommand(message) {
  var lowerMessage = message.toLowerCase();
  
  // Check if this is a REMOVE request (contains "remove" or "un-" prefix)
  var isRemoveRequest = /\b(remove|un-?)\b/.test(lowerMessage);
  
  // Check for REMOVE formatting keywords
  // Handle compound: "remove bold and underline", "remove underline and bold", etc.
  var wantsRemoveBold = isRemoveRequest && /\b(bold)\b/.test(lowerMessage);
  var wantsRemoveItalic = isRemoveRequest && /\b(italic|italics)\b/.test(lowerMessage);
  var wantsRemoveUnderline = isRemoveRequest && /\b(underline|underlined)\b/.test(lowerMessage);
  
  // Also detect specific un- prefixes
  if (/\b(unbold|un-bold)\b/.test(lowerMessage)) wantsRemoveBold = true;
  if (/\b(unitalic|un-italic|unitalicize)\b/.test(lowerMessage)) wantsRemoveItalic = true;
  if (/\b(un-?underline)\b/.test(lowerMessage)) wantsRemoveUnderline = true;
  
  // Check for ADD formatting keywords (only if NOT a remove request)
  var wantsBold = !isRemoveRequest && /\b(bold|make it bold|bold it)\b/.test(lowerMessage);
  var wantsItalic = !isRemoveRequest && /\b(italic|italics|italicize|make it italic)\b/.test(lowerMessage);
  var wantsUnderline = !isRemoveRequest && /\b(underline|underlined)\b/.test(lowerMessage);
  var wantsCenter = /\b(center|centered|centre|centred)\b/.test(lowerMessage);
  var wantsLeft = /\b(left align|align left|left-align)\b/.test(lowerMessage);
  var wantsRight = /\b(right align|align right|right-align)\b/.test(lowerMessage);
  
  // If no formatting intent detected, return false
  var hasFormatIntent = wantsBold || wantsItalic || wantsUnderline || wantsCenter || wantsLeft || wantsRight;
  var hasRemoveIntent = wantsRemoveBold || wantsRemoveItalic || wantsRemoveUnderline;
  
  if (!hasFormatIntent && !hasRemoveIntent) {
    return false;
  }
  
  // Check if user is referring to "first heading", "the title", "the heading", etc.
  var refersToHeading = /\b(first heading|the heading|title|the title|first title|main heading)\b/.test(lowerMessage);
  
  // Check for quoted text like "make 'What is Android?' bold" or 'make "hello" italic'
  var quotedTextMatch = message.match(/['""]([^'""]+)['""]/) || message.match(/'([^']+)'/) || message.match(/"([^"]+)"/);
  var quotedText = quotedTextMatch ? quotedTextMatch[1] : null;
  
  try {
    var applied = [];
    var removed = [];
    var targetDescription = "";
    
    // Case 1: User refers to heading/title
    if (refersToHeading) {
      var formatOptions = {};
      // Add formatting
      if (wantsBold) { formatOptions.bold = true; applied.push("bold"); }
      if (wantsItalic) { formatOptions.italic = true; applied.push("italic"); }
      if (wantsUnderline) { formatOptions.underline = true; applied.push("underlined"); }
      // Remove formatting
      if (wantsRemoveBold) { formatOptions.bold = false; removed.push("bold"); }
      if (wantsRemoveItalic) { formatOptions.italic = false; removed.push("italics"); }
      if (wantsRemoveUnderline) { formatOptions.underline = false; removed.push("underline"); }
      
      if (Object.keys(formatOptions).length > 0) {
        var headingText = await documentService.formatFirstHeading(formatOptions);
        targetDescription = '"' + headingText.substring(0, 30) + (headingText.length > 30 ? '...' : '') + '"';
      }
      
      // Handle alignment for heading
      if (wantsCenter) {
        var headingText2 = await documentService.alignFirstHeading("Center");
        applied.push("centered");
        if (!targetDescription) {
          targetDescription = '"' + headingText2.substring(0, 30) + (headingText2.length > 30 ? '...' : '') + '"';
        }
      } else if (wantsLeft) {
        await documentService.alignFirstHeading("Left");
        applied.push("left-aligned");
      } else if (wantsRight) {
        await documentService.alignFirstHeading("Right");
        applied.push("right-aligned");
      }
      
      if (applied.length > 0 || removed.length > 0) {
        // Build appropriate message
        var msg = "✅ Done! ";
        if (removed.length > 0) {
          msg += "Removed " + removed.join(" and ") + " from " + targetDescription;
        }
        if (applied.length > 0) {
          if (removed.length > 0) msg += " and ";
          msg += "made it " + applied.join(", ");
        }
        msg += ".";
        addAssistantMessage(msg);
        return true;
      }
    }
    
    // Case 2: User specified text in quotes
    if (quotedText) {
      var formatOptions = {};
      // Add formatting
      if (wantsBold) { formatOptions.bold = true; applied.push("bold"); }
      if (wantsItalic) { formatOptions.italic = true; applied.push("italic"); }
      if (wantsUnderline) { formatOptions.underline = true; applied.push("underlined"); }
      // Remove formatting
      if (wantsRemoveBold) { formatOptions.bold = false; removed.push("bold"); }
      if (wantsRemoveItalic) { formatOptions.italic = false; removed.push("italics"); }
      if (wantsRemoveUnderline) { formatOptions.underline = false; removed.push("underline"); }
      
      if (Object.keys(formatOptions).length > 0) {
        var count = await documentService.formatText(quotedText, formatOptions);
        targetDescription = '"' + quotedText + '"';
        var msg = "✅ Done! ";
        if (removed.length > 0) {
          msg += "Removed " + removed.join(" and ") + " from " + targetDescription;
        }
        if (applied.length > 0) {
          if (removed.length > 0) msg += " and ";
          msg += "made it " + applied.join(", ");
        }
        msg += " (" + count + " occurrence" + (count > 1 ? "s" : "") + ").";
        addAssistantMessage(msg);
        return true;
      }
    }
    
    // Case 3: Check if there's selected text (original behavior)
    var hasSelection = await documentService.hasSelection();
    if (!hasSelection) {
      addAssistantMessage("💡 I can format text for you! Try:\n• Select text first, then say \"make it bold\"\n• Or say \"make the first heading bold\"\n• Or say \"make 'specific text' italic\"\n• Or say \"remove bold and underline from heading\"");
      return true;
    }
    
    // Apply formatting to selection
    if (wantsBold || wantsItalic || wantsUnderline || wantsRemoveBold || wantsRemoveItalic || wantsRemoveUnderline) {
      var formatOptions = {};
      if (wantsBold) { formatOptions.bold = true; applied.push("bold"); }
      if (wantsItalic) { formatOptions.italic = true; applied.push("italic"); }
      if (wantsUnderline) { formatOptions.underline = true; applied.push("underlined"); }
      if (wantsRemoveBold) { formatOptions.bold = false; removed.push("bold"); }
      if (wantsRemoveItalic) { formatOptions.italic = false; removed.push("italics"); }
      if (wantsRemoveUnderline) { formatOptions.underline = false; removed.push("underline"); }
      
      await documentService.formatSelection(formatOptions);
    }
    
    // Apply alignment
    if (wantsCenter) {
      await documentService.setAlignment("Center");
      applied.push("centered");
    } else if (wantsLeft) {
      await documentService.setAlignment("Left");
      applied.push("left-aligned");
    } else if (wantsRight) {
      await documentService.setAlignment("Right");
      applied.push("right-aligned");
    }
    
    if (applied.length > 0 || removed.length > 0) {
      var msg = "✅ Done! ";
      if (removed.length > 0) {
        msg += "Removed " + removed.join(" and ") + " from selection";
      }
      if (applied.length > 0) {
        if (removed.length > 0) msg += " and ";
        msg += "made it " + applied.join(", ");
      }
      msg += ".";
      addAssistantMessage(msg);
      return true;
    }
  } catch (error) {
    addAssistantMessage("❌ Error: " + error.message);
    return true;
  }
  
  return false;
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
