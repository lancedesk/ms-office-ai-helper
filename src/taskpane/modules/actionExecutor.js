// Action Executor Module
// Parses AI responses and executes Office.js code dynamically

import { addSystemMessage } from './chatUI.js';

/**
 * Sanitize code extracted from [EXECUTE] block - remove markdown, fix common issues
 * @param {string} code - Raw code from AI response
 * @returns {string} Sanitized code
 */
function sanitizeExecuteCode(code) {
  var s = code.trim();
  s = s.replace(/^```[\w]*\s*/g, '').replace(/\s*```$/g, '').trim();
  s = s.replace(/\bconst\b/g, 'var').replace(/\blet\b/g, 'var');
  return s;
}

/**
 * Validate that code has required structure before execution
 * @param {string} code - Code to validate
 * @returns {boolean} True if code looks valid
 */
function validateExecuteCode(code) {
  if (!code || code.length < 15) return false;
  if (code.indexOf('Word.run') === -1) return false;
  return true;
}

/**
 * Parse AI response and execute any [EXECUTE] code blocks
 * @param {string} response - AI response text
 * @param {object} documentService - Document service instance
 * @returns {Promise<string>} Cleaned response without action markers
 */
async function parseAndExecuteActions(response, documentService) {
  var cleanedResponse = response;
  
  // Handle [EXECUTE] code blocks
  var executeRegex = /\[EXECUTE\]\s*([\s\S]*?)\s*\[\/EXECUTE\]/gi;
  var executeMatch;
  var executedCount = 0;
  var errors = [];
  
  while ((executeMatch = executeRegex.exec(response)) !== null) {
    var code = executeMatch[1].trim();
    
    if (code) {
      try {
        code = sanitizeExecuteCode(code);
        if (!validateExecuteCode(code)) {
          errors.push("Generated code missing required Word.run/context.sync. Try rephrasing your request.");
          continue;
        }
        console.log("Executing dynamic Office.js code:", code.substring(0, 100) + "...");
        
        var asyncFunc = new Function('Word', 'return (async function() { ' + code + ' })()');
        await asyncFunc(Word);
        
        executedCount++;
        console.log("Code executed successfully");
      } catch (error) {
        console.error("Error executing dynamic code:", error);
        var msg = error.message || String(error);
        if (msg.length > 80) msg = msg.substring(0, 77) + "...";
        errors.push(msg);
      }
    }
  }
  
  // Show result message
  if (executedCount > 0) {
    addSystemMessage("✅ Done! Executed " + executedCount + " action" + (executedCount > 1 ? "s" : "") + " successfully.");
  }
  if (errors.length > 0) {
    var errMsg = errors.length === 1 ? errors[0] : errors.join("; ");
    addSystemMessage("⚠️ Action couldn't run: " + errMsg + " Try rephrasing your request.");
  }
  
  // Remove all [EXECUTE] blocks from displayed response
  cleanedResponse = cleanedResponse.replace(/\[EXECUTE\][\s\S]*?\[\/EXECUTE\]/gi, '').trim();
  
  // Handle legacy FORMAT action for backward compatibility
  cleanedResponse = await handleLegacyFormatAction(response, cleanedResponse, documentService);
  
  // Handle legacy INSERT action
  cleanedResponse = await handleLegacyInsertAction(response, cleanedResponse, documentService);
  
  // Handle legacy TABLE action
  cleanedResponse = await handleLegacyTableAction(response, cleanedResponse, documentService);
  
  // Handle legacy CREATE action
  cleanedResponse = await handleLegacyCreateAction(response, cleanedResponse, documentService);
  
  // Handle legacy REPLACE action
  cleanedResponse = await handleLegacyReplaceAction(response, cleanedResponse, documentService);
  
  // Final cleanup
  cleanedResponse = cleanedResponse
    .replace(/\[EXECUTE\][\s\S]*?\[\/EXECUTE\]/gi, '')
    .replace(/\[ACTION:[^\]]*\]/gi, '')
    .replace(/\[\/[A-Z]+\]/gi, '')
    .replace(/\[TOC\]/gi, '')
    .replace(/---CONTENT START---/gi, '')
    .replace(/---CONTENT END---/gi, '')
    .replace(/^\s*[-=]{3,}\s*$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  
  return cleanedResponse;
}

/**
 * Handle legacy FORMAT actions
 */
async function handleLegacyFormatAction(response, cleanedResponse, documentService) {
  var formatRegex = /\[ACTION:\s*FORMAT\s+([^\]]+)\]/gi;
  var formatMatch;
  
  while ((formatMatch = formatRegex.exec(response)) !== null) {
    var actionParams = formatMatch[1];
    var targetMatch = actionParams.match(/target\s*=\s*["']([^"']+)["']/i);
    var target = targetMatch ? targetMatch[1] : null;
    
    var formatOptions = {};
    if (/bold\s*=\s*true/i.test(actionParams)) formatOptions.bold = true;
    else if (/bold\s*=\s*false/i.test(actionParams)) formatOptions.bold = false;
    if (/italic\s*=\s*true/i.test(actionParams)) formatOptions.italic = true;
    else if (/italic\s*=\s*false/i.test(actionParams)) formatOptions.italic = false;
    if (/underline\s*=\s*true/i.test(actionParams)) formatOptions.underline = true;
    else if (/underline\s*=\s*false/i.test(actionParams)) formatOptions.underline = false;
    
    var wantsCenter = /center\s*=\s*true/i.test(actionParams);
    var wantsLeft = /left\s*=\s*true/i.test(actionParams);
    var wantsRight = /right\s*=\s*true/i.test(actionParams);
    
    if (target) {
      try {
        var isHeadingTarget = /first heading|the heading|title|the title/i.test(target);
        
        if (isHeadingTarget) {
          if (Object.keys(formatOptions).length > 0) {
            await documentService.formatFirstHeading(formatOptions);
          }
          if (wantsCenter) await documentService.alignFirstHeading("Center");
          else if (wantsLeft) await documentService.alignFirstHeading("Left");
          else if (wantsRight) await documentService.alignFirstHeading("Right");
        } else {
          if (Object.keys(formatOptions).length > 0) {
            await documentService.formatText(target, formatOptions);
          }
        }
        console.log("FORMAT action executed for target:", target);
      } catch (error) {
        console.error("Error executing FORMAT action:", error);
      }
    }
  }
  
  return cleanedResponse.replace(/\[ACTION:\s*FORMAT\s+[^\]]+\]/gi, '').trim();
}

/**
 * Handle legacy INSERT actions
 * Supports two formats:
 * 1. [ACTION: INSERT heading="X" newpage=true] ---CONTENT START--- ... ---CONTENT END--- (for long content)
 * 2. [ACTION: INSERT heading="X" content="short" newpage=true] (for short content)
 */
async function handleLegacyInsertAction(response, cleanedResponse, documentService) {
  // Format 1: INSERT with ---CONTENT START--- ---CONTENT END--- (best for articles, essays)
  var insertBlockRegex = /\[ACTION:\s*INSERT\s+([^\]]*)\]\s*---CONTENT START---\s*([\s\S]*?)\s*---CONTENT END---/gi;
  var insertBlockMatch = insertBlockRegex.exec(response);
  
  if (insertBlockMatch) {
    var insertParams = insertBlockMatch[1];
    var content = insertBlockMatch[2].trim();
    var headingMatch = insertParams.match(/heading\s*=\s*["']([^"']+)["']/i);
    var heading = headingMatch ? headingMatch[1] : null;
    var newPage = /newpage\s*=\s*true/i.test(insertParams);
    
    if (content && content.length > 10) {
      try {
        content = content.replace(/\[ACTION:[^\]]*\]/gi, '').replace(/\[\/[A-Z]+\]/gi, '').trim();
        await documentService.insertContentSection(heading, content, newPage);
        addSystemMessage("📝 Content inserted into document!");
      } catch (error) {
        console.error("Error executing INSERT action:", error);
        addSystemMessage("⚠️ Failed to insert content: " + error.message);
      }
    }
    cleanedResponse = cleanedResponse.replace(/\[ACTION:\s*INSERT\s+[^\]]*\]\s*---CONTENT START---[\s\S]*?---CONTENT END---/gi, '').trim();
    return cleanedResponse;
  }
  
  // Format 2: INSERT with inline content= (for short content)
  var insertRegex = /\[ACTION:\s*INSERT\s+([^\]]+)\]/gi;
  var insertMatch = insertRegex.exec(response);
  
  if (insertMatch) {
    var insertParams = insertMatch[1];
    var headingMatch = insertParams.match(/heading\s*=\s*["']([^"']+)["']/i);
    var heading = headingMatch ? headingMatch[1] : null;
    var contentMatch = insertParams.match(/content\s*=\s*["']((?:[^"']|\\"|\\')*)["']/i);
    var content = contentMatch ? contentMatch[1].replace(/\\n/g, '\n').replace(/\\"/g, '"').trim() : null;
    var newPage = /newpage\s*=\s*true/i.test(insertParams);
    
    if (content) {
      try {
        await documentService.insertContentSection(heading, content, newPage);
        addSystemMessage("📝 Content inserted into document!");
      } catch (error) {
        console.error("Error executing INSERT action:", error);
        addSystemMessage("⚠️ Failed to insert content: " + error.message);
      }
    }
    cleanedResponse = cleanedResponse.replace(insertRegex, '').trim();
  }
  
  return cleanedResponse;
}

/**
 * Handle legacy TABLE actions
 */
async function handleLegacyTableAction(response, cleanedResponse, documentService) {
  var tableRegex = /\[ACTION:\s*TABLE(?:\s+title\s*=\s*["']([^"']+)["'])?\s*\]\s*([\s\S]*?)\s*\[\/TABLE\]/gi;
  var tableMatch;
  
  while ((tableMatch = tableRegex.exec(response)) !== null) {
    var tableTitle = tableMatch[1] || null;
    var tableContent = tableMatch[2].trim();
    
    if (tableContent) {
      try {
        var lines = tableContent.split('\n').filter(function(line) {
          return line.trim().length > 0;
        });
        
        if (lines.length > 0) {
          var headers = lines[0].split('|').map(function(h) {
            return h.trim();
          }).filter(function(h) {
            return h.length > 0;
          });
          
          var rows = [];
          for (var i = 1; i < lines.length; i++) {
            var row = lines[i].split('|').map(function(cell) {
              return cell.trim();
            }).filter(function(cell) {
              return cell.length > 0 || rows.length === 0;
            });
            if (row.length > 0) rows.push(row);
          }
          
          if (headers.length > 0 && rows.length > 0) {
            await documentService.insertTable(headers, rows, tableTitle);
            addSystemMessage("📊 Table inserted successfully!");
          }
        }
      } catch (error) {
        console.error("Error executing TABLE action:", error);
        addSystemMessage("⚠️ Failed to insert table: " + error.message);
      }
    }
  }
  
  return cleanedResponse.replace(/\[ACTION:\s*TABLE(?:\s+[^\]]+)?\s*\][\s\S]*?\[\/TABLE\]/gi, '').trim();
}

/**
 * Handle legacy CREATE actions
 */
async function handleLegacyCreateAction(response, cleanedResponse, documentService) {
  var createRegex = /\[ACTION:\s*CREATE\s*\]\s*---CONTENT START---\s*([\s\S]*?)\s*---CONTENT END---/gi;
  var createMatch = createRegex.exec(response);
  
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
    newDocContent = newDocContent.replace(/\[ACTION:[^\]]*\]/gi, '');
    newDocContent = newDocContent.replace(/\[\/[A-Z]+\]/gi, '');
    newDocContent = newDocContent.trim();
    
    if (newDocContent && newDocContent.length > 10) {
      try {
        // INSERT into current document instead of creating new (CREATE opens blank doc with no content)
        await documentService.insertContentSection(null, newDocContent, false);
        addSystemMessage("📝 Content inserted into your document!");
      } catch (error) {
        console.error("Error executing CREATE/INSERT action:", error);
        addSystemMessage("⚠️ Failed to insert content: " + error.message);
      }
    }
    cleanedResponse = cleanedResponse.replace(/\[ACTION:\s*CREATE\s*\]\s*---CONTENT START---[\s\S]*/gi, '').trim();
  }
  
  return cleanedResponse;
}

/**
 * Handle legacy REPLACE actions
 */
async function handleLegacyReplaceAction(response, cleanedResponse, documentService) {
  var replaceRegex = /\[ACTION:\s*REPLACE\s*\]\s*---CONTENT START---\s*([\s\S]*?)\s*---CONTENT END---/gi;
  var replaceMatch = replaceRegex.exec(response);
  
  if (!replaceMatch) {
    var lenientRegex = /\[ACTION:\s*REPLACE\s*\]\s*---CONTENT START---\s*([\s\S]+)/gi;
    replaceMatch = lenientRegex.exec(response);
    if (replaceMatch) {
      var content = replaceMatch[1];
      content = content.replace(/\n\n\*\*Note:?\*\*[\s\S]*$/i, '');
      content = content.replace(/\n\nI've (reformatted|corrected|fixed)[\s\S]*$/i, '');
      content = content.replace(/\n\nPlease (let me know|confirm)[\s\S]*$/i, '');
      replaceMatch[1] = content.trim();
    }
  }
  
  if (replaceMatch) {
    var newContent = replaceMatch[1].trim();
    newContent = newContent.replace(/\[ACTION:\s*TABLE[^\]]*\]/gi, '');
    newContent = newContent.replace(/\[\/TABLE\]/gi, '');
    newContent = newContent.replace(/\[ACTION:\s*FORMAT[^\]]*\]/gi, '');
    newContent = newContent.replace(/\[\/ACTION\]/gi, '');
    newContent = newContent.trim();
    
    if (newContent && newContent.length > 50) {
      var confirmed = await showReplaceConfirmation(newContent);
      
      if (confirmed) {
        try {
          await documentService.replaceDocumentContent(newContent);
          addSystemMessage("📝 Document has been reformatted!");
        } catch (error) {
          console.error("Error executing REPLACE action:", error);
          addSystemMessage("⚠️ Failed to apply changes to document.");
        }
      } else {
        addSystemMessage("❌ Document replacement cancelled.");
      }
    }
    cleanedResponse = cleanedResponse.replace(/\[ACTION:\s*REPLACE\s*\]\s*---CONTENT START---[\s\S]*/gi, '').trim();
  }
  
  return cleanedResponse;
}

/**
 * Show confirmation dialog before replacing document content
 */
function showReplaceConfirmation(newContent) {
  return new Promise(function(resolve) {
    var overlay = document.createElement('div');
    overlay.id = 'replace-confirm-overlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:10000;display:flex;align-items:center;justify-content:center;';
    
    var modal = document.createElement('div');
    modal.style.cssText = 'background:white;border-radius:8px;max-width:90%;max-height:80%;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 4px 20px rgba(0,0,0,0.3);';
    
    var header = document.createElement('div');
    header.style.cssText = 'padding:16px;background:#f0f0f0;border-bottom:1px solid #ddd;';
    header.innerHTML = '<h3 style="margin:0;font-size:16px;">⚠️ Review Document Changes</h3><p style="margin:8px 0 0;font-size:12px;color:#666;">Review before applying.</p>';
    
    var preview = document.createElement('div');
    preview.style.cssText = 'padding:16px;overflow-y:auto;max-height:300px;font-family:monospace;font-size:11px;white-space:pre-wrap;background:#fafafa;border-bottom:1px solid #ddd;';
    var previewText = newContent.length > 3000 ? newContent.substring(0, 3000) + '\n\n... [' + (newContent.length - 3000) + ' more characters]' : newContent;
    preview.textContent = previewText;
    
    var stats = document.createElement('div');
    stats.style.cssText = 'padding:8px 16px;background:#e8f4fd;font-size:11px;color:#0066cc;';
    stats.textContent = '📊 ' + newContent.trim().split(/\s+/).length + ' words, ' + newContent.split('\n').length + ' lines';
    
    var buttons = document.createElement('div');
    buttons.style.cssText = 'padding:16px;display:flex;gap:12px;justify-content:flex-end;';
    
    var cancelBtn = document.createElement('button');
    cancelBtn.textContent = '❌ Cancel';
    cancelBtn.style.cssText = 'padding:10px 20px;border:1px solid #ccc;background:#fff;border-radius:4px;cursor:pointer;';
    cancelBtn.onclick = function() { document.body.removeChild(overlay); resolve(false); };
    
    var applyBtn = document.createElement('button');
    applyBtn.textContent = '✅ Apply Changes';
    applyBtn.style.cssText = 'padding:10px 20px;border:none;background:#0078d4;color:white;border-radius:4px;cursor:pointer;';
    applyBtn.onclick = function() { document.body.removeChild(overlay); resolve(true); };
    
    buttons.appendChild(cancelBtn);
    buttons.appendChild(applyBtn);
    modal.appendChild(header);
    modal.appendChild(preview);
    modal.appendChild(stats);
    modal.appendChild(buttons);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    cancelBtn.focus();
  });
}

export {
  parseAndExecuteActions,
  showReplaceConfirmation
};
