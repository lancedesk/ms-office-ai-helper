// Chat UI Module
// Handles message rendering, markdown parsing, and UI helpers

// DOM element references
var chatContainer = null;
var messageInput = null;
var sendButton = null;
var isWelcomeScreen = true;

/**
 * Initialize DOM element references
 */
function initializeElements() {
  if (!chatContainer) {
    chatContainer = document.getElementById("chat-container");
    messageInput = document.getElementById("message-input");
    sendButton = document.getElementById("send-button");
  }
}

/**
 * Get chat container reference
 */
function getChatContainer() {
  initializeElements();
  return chatContainer;
}

/**
 * Get message input reference
 */
function getMessageInput() {
  initializeElements();
  return messageInput;
}

/**
 * Get send button reference
 */
function getSendButton() {
  initializeElements();
  return sendButton;
}

/**
 * Remove welcome screen
 */
function removeWelcomeScreen() {
  if (isWelcomeScreen) {
    var welcomeScreen = document.querySelector(".welcome-screen");
    if (welcomeScreen) {
      welcomeScreen.remove();
    }
    isWelcomeScreen = false;
  }
}

/**
 * Add a system message (grey, centered)
 */
function addSystemMessage(text) {
  initializeElements();
  removeWelcomeScreen();
  
  var messageDiv = document.createElement("div");
  messageDiv.className = "system-message";
  messageDiv.textContent = text;
  chatContainer.appendChild(messageDiv);
  scrollToBottom();
}

/**
 * Add a user message (right-aligned)
 */
function addUserMessage(text) {
  initializeElements();
  removeWelcomeScreen();
  
  var messageDiv = document.createElement("div");
  messageDiv.className = "message user";
  messageDiv.innerHTML = 
    '<div class="message-avatar">U</div>' +
    '<div class="message-content">' + escapeHtml(text) + '</div>';
  chatContainer.appendChild(messageDiv);
  scrollToBottom();
}

/**
 * Add an assistant message (left-aligned, with markdown)
 */
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

/**
 * Show loading indicator
 */
function showLoading() {
  initializeElements();
  
  var loadingDiv = document.createElement("div");
  loadingDiv.className = "message assistant";
  loadingDiv.id = "loading-message";
  loadingDiv.innerHTML = 
    '<div class="message-avatar">AI</div>' +
    '<div class="message-content">' +
      '<div class="loading">' +
        '<div class="loading-dot"></div>' +
        '<div class="loading-dot"></div>' +
        '<div class="loading-dot"></div>' +
      '</div>' +
    '</div>';
  chatContainer.appendChild(loadingDiv);
  scrollToBottom();
}

/**
 * Hide loading indicator
 */
function hideLoading() {
  var loadingMessage = document.getElementById("loading-message");
  if (loadingMessage) {
    loadingMessage.remove();
  }
}

/**
 * Scroll chat to bottom
 */
function scrollToBottom() {
  if (chatContainer) {
    chatContainer.scrollTop = chatContainer.scrollHeight;
  }
}

/**
 * Escape HTML special characters
 */
function escapeHtml(text) {
  var div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

/**
 * Simple markdown parser - converts markdown to HTML
 */
function parseMarkdown(text) {
  if (!text) return '';
  
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
    
    // Unordered list items
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
    
    // Ordered list items
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

// Global function for copy button click
if (typeof window !== 'undefined') {
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
}

export {
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
  scrollToBottom,
  escapeHtml,
  parseMarkdown
};
