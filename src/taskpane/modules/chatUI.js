// Chat UI Module
// Handles message rendering, markdown parsing, and UI helpers

// DOM element references (initialized by initializeElements)
var chatMessages = null;
var messageInput = null;
var sendButton = null;
var loadingIndicator = null;

/**
 * Initialize DOM element references
 */
function initializeElements() {
  chatMessages = document.getElementById("chatMessages");
  messageInput = document.getElementById("messageInput");
  sendButton = document.getElementById("sendButton");
  loadingIndicator = document.getElementById("loadingIndicator");
}

/**
 * Remove welcome screen
 */
function removeWelcomeScreen() {
  var welcomeScreen = document.getElementById('welcomeScreen');
  if (welcomeScreen) {
    welcomeScreen.style.display = 'none';
  }
}

/**
 * Add a system message (grey, centered)
 */
function addSystemMessage(text) {
  initializeElements();
  removeWelcomeScreen();
  var messageDiv = document.createElement("div");
  messageDiv.className = "message system-message";
  messageDiv.innerHTML = '<div class="message-content">' + escapeHtml(text) + '</div>';
  chatMessages.appendChild(messageDiv);
  scrollToBottom();
}

/**
 * Add a user message (right-aligned)
 */
function addUserMessage(text) {
  initializeElements();
  removeWelcomeScreen();
  var messageDiv = document.createElement("div");
  messageDiv.className = "message user-message";
  messageDiv.innerHTML = '<div class="message-content">' + escapeHtml(text) + '</div>';
  chatMessages.appendChild(messageDiv);
  scrollToBottom();
}

/**
 * Add an assistant message (left-aligned, with markdown)
 */
function addAssistantMessage(text) {
  initializeElements();
  removeWelcomeScreen();
  var messageDiv = document.createElement("div");
  messageDiv.className = "message assistant-message";
  
  // Generate unique ID for this message
  var messageId = generateMessageId();
  
  // Parse markdown and render
  var parsedContent = parseMarkdown(text);
  
  messageDiv.innerHTML = 
    '<div class="message-content">' + parsedContent + '</div>' +
    '<div class="message-actions">' +
      '<button class="copy-btn" onclick="copyToClipboard(\'' + messageId + '\', this)" title="Copy to clipboard">' +
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
          '<rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>' +
          '<path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>' +
        '</svg>' +
      '</button>' +
    '</div>';
  
  // Store original text for copying
  messageDiv.dataset.originalText = text;
  messageDiv.dataset.messageId = messageId;
  
  chatMessages.appendChild(messageDiv);
  scrollToBottom();
}

/**
 * Show loading indicator
 */
function showLoading() {
  initializeElements();
  if (loadingIndicator) {
    loadingIndicator.style.display = "flex";
    scrollToBottom();
  }
  if (sendButton) {
    sendButton.disabled = true;
  }
  if (messageInput) {
    messageInput.disabled = true;
  }
}

/**
 * Hide loading indicator
 */
function hideLoading() {
  initializeElements();
  if (loadingIndicator) {
    loadingIndicator.style.display = "none";
  }
  if (sendButton) {
    sendButton.disabled = false;
  }
  if (messageInput) {
    messageInput.disabled = false;
    messageInput.focus();
  }
}

/**
 * Scroll chat to bottom
 */
function scrollToBottom() {
  if (chatMessages) {
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }
}

/**
 * Escape HTML special characters
 */
function escapeHtml(text) {
  var div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

/**
 * Parse markdown to HTML
 */
function parseMarkdown(text) {
  if (!text) return '';
  
  // Escape HTML first
  var html = escapeHtml(text);
  
  // Code blocks (```)
  html = html.replace(/```(\w*)\n([\s\S]*?)```/g, function(match, lang, code) {
    return '<pre class="code-block"><code>' + code.trim() + '</code></pre>';
  });
  
  // Inline code (`)
  html = html.replace(/`([^`]+)`/g, '<code class="inline-code">$1</code>');
  
  // Bold (**text** or __text__)
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/__([^_]+)__/g, '<strong>$1</strong>');
  
  // Italic (*text* or _text_)
  html = html.replace(/\*([^*]+)\*/g, '<em>$1</em>');
  html = html.replace(/_([^_]+)_/g, '<em>$1</em>');
  
  // Headers
  html = html.replace(/^### (.+)$/gm, '<h4>$1</h4>');
  html = html.replace(/^## (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^# (.+)$/gm, '<h2>$1</h2>');
  
  // Bullet lists
  html = html.replace(/^[-*] (.+)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>\n?)+/g, '<ul>$&</ul>');
  
  // Numbered lists
  html = html.replace(/^\d+\. (.+)$/gm, '<li>$1</li>');
  
  // Links
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank">$1</a>');
  
  // Line breaks
  html = html.replace(/\n\n/g, '</p><p>');
  html = html.replace(/\n/g, '<br>');
  
  // Wrap in paragraph
  if (!html.startsWith('<')) {
    html = '<p>' + html + '</p>';
  }
  
  return html;
}

/**
 * Generate unique message ID
 */
function generateMessageId() {
  return 'msg_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

/**
 * Copy message text to clipboard
 */
function copyToClipboard(messageId, buttonElement) {
  var messageDiv = document.querySelector('[data-message-id="' + messageId + '"]');
  if (messageDiv) {
    var text = messageDiv.dataset.originalText;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function() {
        showCopySuccess(buttonElement);
      });
    } else {
      fallbackCopyToClipboard(text, buttonElement);
    }
  }
}

/**
 * Fallback copy method for older browsers
 */
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
    console.error('Copy failed:', err);
  }
  document.body.removeChild(textArea);
}

/**
 * Show copy success feedback
 */
function showCopySuccess(buttonElement) {
  buttonElement.innerHTML = '✓';
  buttonElement.classList.add('copied');
  setTimeout(function() {
    buttonElement.innerHTML = 
      '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
        '<rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>' +
        '<path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>' +
      '</svg>';
    buttonElement.classList.remove('copied');
  }, 2000);
}

// Make copyToClipboard available globally for onclick handlers
window.copyToClipboard = function(messageId, buttonElement) {
  copyToClipboard(messageId, buttonElement);
};

// Export functions
export {
  initializeElements,
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
