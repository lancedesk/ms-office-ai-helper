// Action Executor Module
// Parses AI responses and executes Office.js code dynamically

import { addSystemMessage } from './chatUI.js';

/**
 * Parse AI response and execute any [EXECUTE] code blocks
 * Also handles legacy [ACTION:...] formats for backward compatibility
 * @param {string} response - AI response text
 * @param {object} documentService - Document service instance
 * @returns {Promise<string>} Cleaned response without action markers
 */
async function parseAndExecuteActions(response, documentService) {
  var cleanedResponse = response;
  
  // Handle [EXECUTE] code blocks - the primary dynamic approach
  var executeRegex = /\[EXECUTE\]\s*([\s\S]*?)\s*\[\/EXECUTE\]/gi;
  var executeMatch;
  var executedCount = 0;
  var errors = [];
  
  while ((executeMatch = executeRegex.exec(response)) !== null) {
    var code = executeMatch[1].trim();
    
    if (code) {
      try {
        console.log("Executing dynamic Office.js code:", code.substring(0, 100) + "...");
        
        // Execute the code dynamically
        // The code should be an async function body that uses Word.run
        var asyncFunc = new Function('Word', 'context', 'return (async () => { ' + code + ' })()');
        await asyncFunc(Word, null);
        
        executedCount++;
        console.log("Code executed successfully");
      } catch (error) {
        console.error("Error executing dynamic code:", error);
        errors.push(error.message);
      }
    }
  }
  
  // Show result message
  if (executedCount > 0) {
    addSystemMessage("✅ Done! Executed " + executedCount + " action" + (executedCount > 1 ? "s" : "") + " successfully.");
  }
  if (errors.length > 0) {
    addSystemMessage("⚠️ Some actions failed: " + errors.join(", "));
  }
  
  // Remove all [EXECUTE] blocks from displayed response
  cleanedResponse = cleanedResponse.replace(/\[EXECUTE\][\s\S]*?\[\/EXECUTE\]/gi, '').trim();
  
  // Final cleanup: remove ANY remaining action-like patterns that users shouldn't see
  cleanedResponse = cleanedResponse
    .replace(/\[EXECUTE\][\s\S]*?\[\/EXECUTE\]/gi, '')  // [EXECUTE]...[/EXECUTE]
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
    
    // Buttons
    var buttons = document.createElement('div');
    buttons.style.cssText = 'padding:16px;display:flex;gap:12px;justify-content:flex-end;';
    
    var cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.style.cssText = 'padding:8px 16px;border:1px solid #ccc;border-radius:4px;background:white;cursor:pointer;';
    cancelBtn.onclick = function() {
      document.body.removeChild(overlay);
      resolve(false);
    };
    
    var confirmBtn = document.createElement('button');
    confirmBtn.textContent = 'Apply Changes';
    confirmBtn.style.cssText = 'padding:8px 16px;border:none;border-radius:4px;background:#0078d4;color:white;cursor:pointer;';
    confirmBtn.onclick = function() {
      document.body.removeChild(overlay);
      resolve(true);
    };
    
    buttons.appendChild(cancelBtn);
    buttons.appendChild(confirmBtn);
    
    modal.appendChild(header);
    modal.appendChild(preview);
    modal.appendChild(buttons);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
  });
}

// Export functions
export {
  parseAndExecuteActions,
  showReplaceConfirmation
};
