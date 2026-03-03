// Document Service
// Handles all Word document operations (reading, writing, formatting)

class DocumentService {
  constructor() {
    this.maxContextLength = 10000; // Characters to include in context
  }

  /**
   * Read the entire document body text
   * @returns {Promise<string>} Document text content
   */
  async readDocumentText() {
    return new Promise((resolve, reject) => {
      if (typeof Word === "undefined") {
        reject(new ReferenceError("Word is not defined. Ensure this is running inside the Word client."));
        return;
      }

      Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        resolve(body.text);
      }).catch(reject);
    });
  }

  /**
   * Get document with structured content (paragraphs)
   * @returns {Promise<Object>} Document structure with paragraphs
   */
  async readDocumentStructure() {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        
        paragraphs.load("text,style");
        await context.sync();

        const structure = {
          paragraphs: [],
          totalParagraphs: paragraphs.items.length
        };

        paragraphs.items.forEach((paragraph, index) => {
          structure.paragraphs.push({
            index: index,
            text: paragraph.text,
            style: paragraph.style,
            isEmpty: paragraph.text.trim().length === 0
          });
        });

        resolve(structure);
      }).catch(reject);
    });
  }

  /**
   * Get document metadata
   * @returns {Promise<Object>} Document properties
   */
  async getDocumentMetadata() {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const properties = context.document.properties;
        const body = context.document.body;
        
        properties.load("title,author,subject,keywords,comments,lastAuthor,creationDate,lastPrintDate");
        body.load("text");
        
        await context.sync();

        const text = body.text;
        const words = text.trim().split(/\s+/).filter(w => w.length > 0);
        const characters = text.length;
        const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 0);

        resolve({
          title: properties.title || "Untitled Document",
          author: properties.author || "Unknown",
          subject: properties.subject || "",
          keywords: properties.keywords || "",
          lastAuthor: properties.lastAuthor || "",
          creationDate: properties.creationDate,
          lastPrintDate: properties.lastPrintDate,
          statistics: {
            wordCount: words.length,
            characterCount: characters,
            characterCountNoSpaces: text.replace(/\s/g, '').length,
            sentenceCount: sentences.length,
            paragraphCount: text.split(/\n\n+/).length
          }
        });
      }).catch(reject);
    });
  }

  /**
   * Get document content optimized for AI context
   * Chunks large documents and includes metadata
   * @param {number} maxLength - Maximum characters to return
   * @returns {Promise<Object>} Formatted context object
   */
  async getDocumentContext(maxLength = null) {
    maxLength = maxLength || this.maxContextLength;

    try {
      const [text, metadata, structure] = await Promise.all([
        this.readDocumentText(),
        this.getDocumentMetadata(),
        this.readDocumentStructure()
      ]);

      // Check if document is empty
      if (!text || text.trim().length === 0) {
        return {
          isEmpty: true,
          message: "Document is empty"
        };
      }

      // Prepare context
      let context = {
        isEmpty: false,
        metadata: metadata,
        content: text,
        contentLength: text.length,
        isTruncated: false,
        structure: {
          paragraphCount: structure.totalParagraphs,
          hasHeaders: structure.paragraphs.some(p => p.style.includes("Heading"))
        }
      };

      // Truncate if needed
      if (text.length > maxLength) {
        context.content = text.substring(0, maxLength);
        context.isTruncated = true;
        context.truncatedAt = maxLength;
      }

      return context;
    } catch (error) {
      console.error("Error getting document context:", error);
      throw error;
    }
  }

  /**
   * Format document context for AI prompt
   * @param {Object} context - Document context object
   * @returns {string} Formatted string for AI
   */
  formatContextForAI(context) {
    if (context.isEmpty) {
      return "The document is currently empty. When the user asks to write content, use [ACTION: INSERT] to add it here.";
    }

    let formatted = `Document Information:\n`;
    formatted += `- Title: ${context.metadata.title}\n`;
    formatted += `- Word Count: ${context.metadata.statistics.wordCount}\n`;
    formatted += `- Paragraph Count: ${context.structure.paragraphCount}\n`;
    
    if (context.metadata.author) {
      formatted += `- Author: ${context.metadata.author}\n`;
    }

    formatted += `\nDocument Content:\n`;
    formatted += context.content;
    formatted += `\n\nIMPORTANT: The user has content in this document. When they ask to write something, use [ACTION: INSERT] to APPEND at the end. Do NOT create a new document.`;

    if (context.isTruncated) {
      formatted += `\n\n[Content truncated at ${context.truncatedAt} characters. Full document has ${context.contentLength} characters.]`;
    }

    return formatted;
  }

  /**
   * Insert text at current cursor position
   * @param {string} text - Text to insert
   * @param {string} location - Where to insert (default: "Replace")
   * @returns {Promise<void>}
   */
  async insertText(text, location = "Replace") {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(text, location);
        await context.sync();
        resolve();
      }).catch(reject);
    });
  }

  /**
   * Replace selected text
   * @param {string} newText - New text to replace selection
   * @returns {Promise<void>}
   */
  async replaceSelection(newText) {
    return this.insertText(newText, "Replace");
  }

  /**
   * Insert text at the end of document
   * @param {string} text - Text to append
   * @returns {Promise<void>}
   */
  async appendText(text) {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const body = context.document.body;
        body.insertText(text, Word.InsertLocation.end);
        await context.sync();
        resolve();
      }).catch(reject);
    });
  }

  /**
   * Get selected text
   * @returns {Promise<string>} Selected text or empty string
   */
  async getSelectedText() {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("text");
        await context.sync();
        resolve(range.text);
      }).catch(reject);
    });
  }

  /**
   * Check if there is a selection
   * @returns {Promise<boolean>} True if text is selected
   */
  async hasSelection() {
    try {
      const selected = await this.getSelectedText();
      return selected && selected.trim().length > 0;
    } catch {
      return false;
    }
  }

  /**
   * Search for text in document
   * @param {string} searchText - Text to search for
   * @returns {Promise<Array>} Array of search results
   */
  async searchDocument(searchText) {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        const results = context.document.body.search(searchText, { matchCase: false });
        results.load("text");
        await context.sync();
        
        const matches = results.items.map((item, index) => ({
          index: index,
          text: item.text
        }));
        
        resolve(matches);
      }).catch(reject);
    });
  }

  /**
   * Clear all document content
   * @returns {Promise<void>}
   */
  async clearDocument() {
    return new Promise((resolve, reject) => {
      Word.run(async (context) => {
        context.document.body.clear();
        await context.sync();
        resolve();
      }).catch(reject);
    });
  }

  /**
   * Format selected text with specified styles
   * @param {Object} options - Formatting options (bold, italic, underline, color, alignment)
   * @returns {Promise<void>}
   */
  async formatSelection(options) {
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var range = context.document.getSelection();
        range.load("text");
        
        return context.sync().then(function() {
          if (!range.text || range.text.trim().length === 0) {
            reject(new Error("No text selected"));
            return;
          }
          
          var font = range.font;
          
          if (options.bold !== undefined) {
            font.bold = options.bold;
          }
          if (options.italic !== undefined) {
            font.italic = options.italic;
          }
          if (options.underline !== undefined) {
            font.underline = options.underline ? "Single" : "None";
          }
          if (options.color) {
            font.color = options.color;
          }
          if (options.size) {
            font.size = options.size;
          }
          if (options.highlightColor) {
            font.highlightColor = options.highlightColor;
          }
          
          return context.sync();
        }).then(function() {
          resolve();
        });
      }).catch(reject);
    });
  }

  /**
   * Set paragraph alignment for selection
   * @param {string} alignment - "Left", "Center", "Right", "Justified"
   * @returns {Promise<void>}
   */
  async setAlignment(alignment) {
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var range = context.document.getSelection();
        var paragraphs = range.paragraphs;
        paragraphs.load("items");
        
        return context.sync().then(function() {
          for (var i = 0; i < paragraphs.items.length; i++) {
            paragraphs.items[i].alignment = alignment;
          }
          return context.sync();
        }).then(function() {
          resolve();
        });
      }).catch(reject);
    });
  }

  /**
   * Apply heading style to selection
   * @param {number} level - Heading level (1-6)
   * @returns {Promise<void>}
   */
  async applyHeading(level) {
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var range = context.document.getSelection();
        var paragraphs = range.paragraphs;
        paragraphs.load("items");
        
        return context.sync().then(function() {
          var styleName = level === 0 ? "Normal" : "Heading " + level;
          for (var i = 0; i < paragraphs.items.length; i++) {
            paragraphs.items[i].style = styleName;
          }
          return context.sync();
        }).then(function() {
          resolve();
        });
      }).catch(reject);
    });
  }

  /**
   * Search and replace text in document
   * @param {string} searchText - Text to find
   * @param {string} replaceText - Text to replace with
   * @returns {Promise<number>} Number of replacements made
   */
  async searchAndReplace(searchText, replaceText) {
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var results = context.document.body.search(searchText, { matchCase: false });
        results.load("items");
        
        return context.sync().then(function() {
          var count = results.items.length;
          for (var i = 0; i < results.items.length; i++) {
            results.items[i].insertText(replaceText, "Replace");
          }
          return context.sync().then(function() {
            resolve(count);
          });
        });
      }).catch(reject);
    });
  }

  /**
   * Find text and apply formatting to it (no selection needed)
   * @param {string} searchText - Text to find and format
   * @param {Object} options - Formatting options (bold, italic, underline, etc.)
   * @returns {Promise<number>} Number of matches formatted
   */
  async formatText(searchText, options) {
    var self = this;
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var results = context.document.body.search(searchText, { matchCase: false });
        results.load("items");
        
        return context.sync().then(function() {
          if (results.items.length === 0) {
            reject(new Error("Text not found: \"" + searchText + "\""));
            return;
          }
          
          for (var i = 0; i < results.items.length; i++) {
            var font = results.items[i].font;
            if (options.bold !== undefined) font.bold = options.bold;
            if (options.italic !== undefined) font.italic = options.italic;
            if (options.underline !== undefined) font.underline = options.underline ? "Single" : "None";
            if (options.color) font.color = options.color;
            if (options.size) font.size = options.size;
            if (options.highlightColor) font.highlightColor = options.highlightColor;
          }
          
          return context.sync().then(function() {
            resolve(results.items.length);
          });
        });
      }).catch(reject);
    });
  }

  /**
   * Get the first heading in the document
   * @returns {Promise<string|null>} First heading text or null
   */
  async getFirstHeading() {
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var paragraphs = context.document.body.paragraphs;
        paragraphs.load("items,text,style");
        
        return context.sync().then(function() {
          for (var i = 0; i < paragraphs.items.length; i++) {
            var style = paragraphs.items[i].style || "";
            if (style.indexOf("Heading") !== -1 || style.indexOf("Title") !== -1) {
              resolve(paragraphs.items[i].text.trim());
              return;
            }
          }
          // No heading found, return first non-empty paragraph
          for (var j = 0; j < paragraphs.items.length; j++) {
            var text = paragraphs.items[j].text.trim();
            if (text.length > 0) {
              resolve(text);
              return;
            }
          }
          resolve(null);
        });
      }).catch(reject);
    });
  }

  /**
   * Format the first heading/title in the document
   * @param {Object} options - Formatting options
   * @returns {Promise<string>} The heading text that was formatted
   */
  async formatFirstHeading(options) {
    var self = this;
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var paragraphs = context.document.body.paragraphs;
        paragraphs.load("items,text,style");
        
        return context.sync().then(function() {
          var targetPara = null;
          
          // Find first heading or title
          for (var i = 0; i < paragraphs.items.length; i++) {
            var style = paragraphs.items[i].style || "";
            if (style.indexOf("Heading") !== -1 || style.indexOf("Title") !== -1) {
              targetPara = paragraphs.items[i];
              break;
            }
          }
          
          // If no heading, use first non-empty paragraph
          if (!targetPara) {
            for (var j = 0; j < paragraphs.items.length; j++) {
              if (paragraphs.items[j].text.trim().length > 0) {
                targetPara = paragraphs.items[j];
                break;
              }
            }
          }
          
          if (!targetPara) {
            reject(new Error("No heading or text found in document"));
            return;
          }
          
          var font = targetPara.font;
          if (options.bold !== undefined) font.bold = options.bold;
          if (options.italic !== undefined) font.italic = options.italic;
          if (options.underline !== undefined) font.underline = options.underline ? "Single" : "None";
          if (options.color) font.color = options.color;
          
          var headingText = targetPara.text.trim();
          
          return context.sync().then(function() {
            resolve(headingText);
          });
        });
      }).catch(reject);
    });
  }

  /**
   * Format first heading with alignment
   * @param {string} alignment - "Left", "Center", "Right", "Justified"
   * @returns {Promise<string>} The heading text that was formatted
   */
  async alignFirstHeading(alignment) {
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var paragraphs = context.document.body.paragraphs;
        paragraphs.load("items,text,style");
        
        return context.sync().then(function() {
          var targetPara = null;
          
          // Find first heading or title
          for (var i = 0; i < paragraphs.items.length; i++) {
            var style = paragraphs.items[i].style || "";
            if (style.indexOf("Heading") !== -1 || style.indexOf("Title") !== -1) {
              targetPara = paragraphs.items[i];
              break;
            }
          }
          
          // If no heading, use first non-empty paragraph
          if (!targetPara) {
            for (var j = 0; j < paragraphs.items.length; j++) {
              if (paragraphs.items[j].text.trim().length > 0) {
                targetPara = paragraphs.items[j];
                break;
              }
            }
          }
          
          if (!targetPara) {
            reject(new Error("No heading or text found in document"));
            return;
          }
          
          targetPara.alignment = alignment;
          var headingText = targetPara.text.trim();
          
          return context.sync().then(function() {
            resolve(headingText);
          });
        });
      }).catch(reject);
    });
  }

  /**
   * Analyze document and provide insights
   * @returns {Promise<Object>} Document analysis
   */
  async analyzeDocument() {
    try {
      const context = await this.getDocumentContext();
      
      if (context.isEmpty) {
        return {
          isEmpty: true,
          message: "Document is empty, nothing to analyze"
        };
      }

      const text = context.content;
      const words = text.trim().split(/\s+/).filter(w => w.length > 0);
      const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 0);
      const paragraphs = text.split(/\n\n+/).filter(p => p.trim().length > 0);

      // Calculate reading time (average 200 words per minute)
      const readingTime = Math.ceil(words.length / 200);

      // Get word frequency
      const wordFreq = {};
      words.forEach(word => {
        const lower = word.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (lower.length > 3) { // Only count words longer than 3 chars
          wordFreq[lower] = (wordFreq[lower] || 0) + 1;
        }
      });

      // Get top 10 most frequent words
      const topWords = Object.entries(wordFreq)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([word, count]) => ({ word, count }));

      return {
        isEmpty: false,
        statistics: context.metadata.statistics,
        readingTime: readingTime,
        averageWordsPerSentence: Math.round(words.length / sentences.length),
        averageWordsPerParagraph: Math.round(words.length / paragraphs.length),
        topWords: topWords,
        hasHeaders: context.structure.hasHeaders,
        metadata: context.metadata
      };
    } catch (error) {
      console.error("Error analyzing document:", error);
      throw error;
    }
  }

  /**
   * Insert text at the end of the document
   * @param {string} text - The text content to insert
   * @param {boolean} addPageBreak - Whether to add a page break before the content
   * @returns {Promise<boolean>} Success status
   */
  async insertTextAtEnd(text, addPageBreak) {
    if (addPageBreak === undefined) addPageBreak = false;
    
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var body = context.document.body;
        
        // Add page break if requested
        if (addPageBreak) {
          body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
        }
        
        // Insert the text at the end
        body.insertText(text, Word.InsertLocation.end);
        
        return context.sync().then(function() {
          resolve(true);
        });
      }).catch(function(error) {
        console.error("Error inserting text:", error);
        reject(error);
      });
    });
  }

  /**
   * Insert content with optional heading at end of document
   * @param {string} heading - Optional heading text
   * @param {string} content - The content to insert
   * @param {boolean} newPage - Whether to start on a new page
   * @returns {Promise<boolean>} Success status
   */
  async insertContentSection(heading, content, newPage) {
    if (newPage === undefined) newPage = false;
    
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var body = context.document.body;
        
        // Add page break for new section (when newPage=true)
        if (newPage) {
          body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
        } else {
          // Add blank paragraph for visual separation when appending
          body.insertParagraph('', Word.InsertLocation.end);
        }
        
        // Add heading if provided
        if (heading) {
          var headingPara = body.insertParagraph(heading, Word.InsertLocation.end);
          headingPara.styleBuiltIn = Word.Style.heading1;
        }
        
        // Add content paragraphs
        var paragraphs = content.split('\n\n');
        paragraphs.forEach(function(para) {
          if (para.trim()) {
            body.insertParagraph(para.trim(), Word.InsertLocation.end);
          }
        });
        
        return context.sync().then(function() {
          resolve(true);
        });
      }).catch(function(error) {
        console.error("Error inserting content section:", error);
        reject(error);
      });
    });
  }

  /**
   * Replace entire document content with properly formatted content
   * Handles headings (lines starting with # or **), bullet points, tables, etc.
   * @param {string} content - The new content with markdown-like formatting
   * @returns {Promise<boolean>} Success status
   */
  async replaceDocumentContent(content) {
    var self = this;
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var body = context.document.body;
        
        // Clear existing content
        body.clear();
        
        // Split content into lines
        var lines = content.split('\n');
        var i = 0;
        
        // Helper to strip markdown bold/italic
        function stripMarkdown(text) {
          return text
            .replace(/\*\*\*(.+?)\*\*\*/g, '$1')  // ***bold italic***
            .replace(/\*\*(.+?)\*\*/g, '$1')       // **bold**
            .replace(/\*(.+?)\*/g, '$1')           // *italic*
            .replace(/__(.+?)__/g, '$1')           // __underline__
            .replace(/_(.+?)_/g, '$1')             // _italic_
            .trim();
        }
        
        // Helper to detect if line is a table row
        function isTableRow(line) {
          // Must have at least 2 pipe characters and content between them
          var pipeCount = (line.match(/\|/g) || []).length;
          return pipeCount >= 2 && !/^\s*\|?\s*[-:]+\s*\|/.test(line); // Not a separator row
        }
        
        // Helper to detect table separator row (| --- | --- |)
        function isTableSeparator(line) {
          return /^\s*\|?\s*[-:]+\s*\|/.test(line) && /[-:]{3,}/.test(line);
        }
        
        while (i < lines.length) {
          var line = lines[i].trim();
          
          // Skip empty lines
          if (!line) {
            i++;
            continue;
          }
          
          // Skip table separator rows
          if (isTableSeparator(line)) {
            i++;
            continue;
          }
          
          // Check if this starts a table (line has | characters)
          if (isTableRow(line)) {
            // Collect all table rows
            var tableRows = [];
            while (i < lines.length) {
              var tableLine = lines[i].trim();
              if (!tableLine) {
                i++;
                break;
              }
              if (isTableSeparator(tableLine)) {
                i++;
                continue;
              }
              if (isTableRow(tableLine)) {
                // Parse cells
                var cells = tableLine.split('|')
                  .map(function(c) { return stripMarkdown(c.trim()); })
                  .filter(function(c) { return c.length > 0; });
                if (cells.length > 0) {
                  tableRows.push(cells);
                }
                i++;
              } else {
                break;
              }
            }
            
            // Create Word table if we have rows
            if (tableRows.length > 1) {
              var numCols = Math.max.apply(null, tableRows.map(function(r) { return r.length; }));
              var numRows = tableRows.length;
              
              var table = body.insertTable(numRows, numCols, Word.InsertLocation.end, null);
              
              for (var row = 0; row < tableRows.length; row++) {
                for (var col = 0; col < tableRows[row].length && col < numCols; col++) {
                  var cell = table.getCell(row, col);
                  cell.value = tableRows[row][col];
                  // Bold the header row
                  if (row === 0) {
                    cell.body.font.bold = true;
                    cell.shadingColor = '#E0E0E0';
                  }
                }
              }
              
              try {
                table.styleBuiltIn = Word.Style.gridTable4_Accent1;
              } catch (e) {
                // Style might not be available
              }
            }
            continue;
          }
          
          // Check for heading patterns
          // **Bold Text** on its own line = Heading
          var isBoldHeading = /^\*\*[^*]+\*\*\s*$/.test(line);
          var isHashHeading1 = /^#\s+/.test(line);
          var isHashHeading2 = /^##\s+/.test(line);
          var isHashHeading3 = /^###\s+/.test(line);
          var isBullet = /^[•\-\*\+]\s+/.test(line) && !/^\*\*/.test(line);
          var isNumbered = /^\d+\.\s+/.test(line);
          
          if (isHashHeading3) {
            var headingText = stripMarkdown(line.replace(/^###\s+/, ''));
            var para = body.insertParagraph(headingText, Word.InsertLocation.end);
            para.styleBuiltIn = Word.Style.heading3;
          } else if (isHashHeading2) {
            var headingText = stripMarkdown(line.replace(/^##\s+/, ''));
            var para = body.insertParagraph(headingText, Word.InsertLocation.end);
            para.styleBuiltIn = Word.Style.heading2;
          } else if (isHashHeading1) {
            var headingText = stripMarkdown(line.replace(/^#\s+/, ''));
            var para = body.insertParagraph(headingText, Word.InsertLocation.end);
            para.styleBuiltIn = Word.Style.heading1;
          } else if (isBoldHeading) {
            // **Heading Text** becomes Heading 2
            var headingText = stripMarkdown(line);
            var para = body.insertParagraph(headingText, Word.InsertLocation.end);
            para.styleBuiltIn = Word.Style.heading2;
          } else if (isBullet) {
            var bulletText = stripMarkdown(line.replace(/^[•\-\*\+]\s+/, ''));
            var para = body.insertParagraph(bulletText, Word.InsertLocation.end);
            para.styleBuiltIn = Word.Style.listBullet;
          } else if (isNumbered) {
            var numberedText = stripMarkdown(line.replace(/^\d+\.\s+/, ''));
            var para = body.insertParagraph(numberedText, Word.InsertLocation.end);
            para.styleBuiltIn = Word.Style.listNumber;
          } else {
            // Regular paragraph
            var paragraphText = stripMarkdown(line);
            body.insertParagraph(paragraphText, Word.InsertLocation.end);
          }
          
          i++;
        }
        
        return context.sync().then(function() {
          resolve(true);
        });
      }).catch(function(error) {
        console.error("Error replacing document content:", error);
        reject(error);
      });
    });
  }

  /**
   * Insert a table into the document
   * @param {Array<string>} headers - Array of column headers
   * @param {Array<Array<string>>} rows - 2D array of row data
   * @param {string} title - Optional title above the table
   * @returns {Promise<boolean>} Success status
   */
  async insertTable(headers, rows, title) {
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        var body = context.document.body;
        
        // Add title if provided
        if (title) {
          var titlePara = body.insertParagraph(title, Word.InsertLocation.end);
          titlePara.styleBuiltIn = Word.Style.heading2;
        }
        
        // Create table with headers + data rows
        var totalRows = rows.length + 1; // +1 for header row
        var totalCols = headers.length;
        
        // Insert table at end of document
        var table = body.insertTable(totalRows, totalCols, Word.InsertLocation.end, null);
        
        // Set header row values
        for (var col = 0; col < headers.length; col++) {
          var cell = table.getCell(0, col);
          cell.value = headers[col];
          cell.body.font.bold = true;
          cell.shadingColor = "#E0E0E0"; // Light gray background for headers
        }
        
        // Set data row values
        for (var row = 0; row < rows.length; row++) {
          for (var col = 0; col < rows[row].length && col < totalCols; col++) {
            var cell = table.getCell(row + 1, col);
            cell.value = rows[row][col] || '';
          }
        }
        
        // Add some basic styling
        table.styleBuiltIn = Word.Style.gridTable4_Accent1;
        
        return context.sync().then(function() {
          console.log("Table inserted successfully:", totalRows, "rows x", totalCols, "cols");
          resolve(true);
        });
      }).catch(function(error) {
        console.error("Error inserting table:", error);
        reject(error);
      });
    });
  }
  
  /**
   * Create a new document with optional content
   * Note: This uses Word.Application.createDocument which works on desktop Word
   * For Word Online, it may have limitations
   * @param {string} content - Optional content to add to the new document (markdown format)
   * @returns {Promise<boolean>} True if successful
   */
  async createDocument(content) {
    var self = this;
    return new Promise(function(resolve, reject) {
      Word.run(function(context) {
        // Try to create a new document using Application.createDocument
        // This creates a blank document in a new window
        try {
          var app = context.application;
          var newDoc = app.createDocument();
          newDoc.load();
          
          return context.sync().then(function() {
            // Open the new document
            newDoc.open();
            return context.sync();
          }).then(function() {
            // If content was provided, we need to add it to the new document
            // But since it's a new window, we'll add a message
            console.log("New document created successfully");
            resolve({ success: true, hasContent: !!content, contentToAdd: content });
          });
        } catch (e) {
          // Fallback: createDocument may not be available in all contexts
          // In that case, offer to clear current document instead
          console.warn("createDocument not available, using fallback:", e.message);
          reject(new Error("CREATE_NOT_SUPPORTED"));
        }
      }).catch(function(error) {
        console.error("Error creating document:", error);
        if (error.message && error.message.includes("not supported")) {
          reject(new Error("CREATE_NOT_SUPPORTED"));
        } else {
          reject(error);
        }
      });
    });
  }
}

// Export for use in other files
export default DocumentService;
