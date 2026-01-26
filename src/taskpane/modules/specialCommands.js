// Special Commands Module
// Handles slash commands like /bold, /analyze, /help, etc.

import { addAssistantMessage, showLoading, hideLoading } from './chatUI.js';

// Document service reference - set by main module
let documentService = null;

/**
 * Initialize with document service
 */
function initSpecialCommands(deps) {
  documentService = deps.documentService;
}

/**
 * Handle special commands (starting with /)
 * @param {string} message - User message
 * @returns {Promise<boolean>} True if command was handled
 */
async function handleSpecialCommand(message) {
  var command = message.toLowerCase().trim();
  
  if (command === '/analyze' || command === '/stats') {
    showLoading();
    try {
      var analysis = await documentService.analyzeDocument();
      hideLoading();
      
      if (analysis.isEmpty) {
        addAssistantMessage("📄 The document is empty. Add some content first!");
      } else {
        var response = "📊 **Document Analysis**\n\n";
        response += "**Statistics:**\n";
        response += "- Words: " + analysis.statistics.wordCount + "\n";
        response += "- Characters: " + analysis.statistics.characterCount + "\n";
        response += "- Sentences: " + analysis.statistics.sentenceCount + "\n";
        response += "- Paragraphs: " + analysis.statistics.paragraphCount + "\n";
        response += "- Reading Time: ~" + analysis.readingTime + " minute(s)\n";
        response += "- Avg Words/Sentence: " + analysis.averageWordsPerSentence + "\n\n";
        
        if (analysis.topWords && analysis.topWords.length > 0) {
          response += "**Most Common Words:**\n";
          analysis.topWords.slice(0, 5).forEach(function(item, i) {
            response += (i + 1) + ". \"" + item.word + "\" (" + item.count + "x)\n";
          });
        }
        
        addAssistantMessage(response);
      }
      return true;
    } catch (error) {
      hideLoading();
      addAssistantMessage("❌ Error analyzing document: " + error.message);
      return true;
    }
  }
  
  if (command === '/help') {
    var helpText = "🤖 **AI Helper Commands**\n\n" +
      "**Slash Commands (select text first):**\n" +
      "• /bold, /italic, /underline\n" +
      "• /center, /left, /right\n" +
      "• /h1, /h2, /h3 - Heading styles\n" +
      "• /analyze - Document stats\n\n" +
      "**Natural Language (no selection needed!):**\n" +
      "• \"make the first heading bold\"\n" +
      "• \"center the title\"\n" +
      "• \"make 'What is Android?' italic\"\n" +
      "• \"underline the heading\"\n\n" +
      "**With Selection:**\n" +
      "• \"make it bold and italic\"\n" +
      "• \"center this\"\n\n" +
      "💡 Just describe what you want - I'll do it!";
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
    if (error.message.indexOf("No text selected") >= 0) {
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

export {
  initSpecialCommands,
  handleSpecialCommand,
  applyFormatting,
  applyAlignment,
  applyHeadingStyle
};
