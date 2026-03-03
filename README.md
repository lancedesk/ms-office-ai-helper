# MS Office AI Helper

An intelligent AI-powered assistant add-in for Microsoft Word that integrates with Groq and Google Gemini AI. Write articles, edit documents, and perform any document action using natural language—in any language.

## ✨ Features

- 📝 **Write Content** - Generate and insert articles, essays, assignments, and reports directly into your document with proper formatting
- ✍️ **Smart Editing** - Format, search/replace, restructure, or rewrite content with plain-English commands
- 🗣️ **Natural Language** - Ask in any language: "delete the doc", "move this word to the top", "swap paragraphs", "center the title"—the AI figures out the right Word APIs
- 📄 **Document Analysis** - Read, summarize, and understand your document content
- 💡 **Summaries** - Get quick summaries of your documents
- 🎨 **Formatting** - Create headers, tables, and more with simple commands
- 🔄 **Multi-Provider Support** - Switch between Groq (Llama 3.3 70B) and Google Gemini AI
- 📐 **Auto-Expanding Input** - Chat input grows as you type longer messages

## 🌐 Live Demo

| Resource | URL |
|----------|-----|
| **Add-in Home** | [https://lancedesk.github.io/ms-office-ai-helper/](https://lancedesk.github.io/ms-office-ai-helper/) |
| **Manifest File** | [https://lancedesk.github.io/ms-office-ai-helper/manifest.xml](https://lancedesk.github.io/ms-office-ai-helper/manifest.xml) |

## 📋 Prerequisites

- Microsoft Word 2016 or later (Windows/Mac) or Word Online
- Node.js 18+ and npm (for development only)
- A Groq API key ([Get one here](https://console.groq.com/keys)) or Google Gemini API key ([Get one here](https://aistudio.google.com/apikey))

## 🚀 Quick Start

### Option 1: Use the Hosted Version (Recommended)

1. **Download the manifest:** [manifest.xml](https://lancedesk.github.io/ms-office-ai-helper/manifest.xml)
2. Open Word → **Insert** → **Get Add-ins** → **My Add-ins** → **Upload My Add-in**
3. Upload the downloaded manifest.xml
4. Enter your API key when prompted (Settings ⚙️)

### Option 2: Development Setup

```bash
# Clone the repository
git clone https://github.com/lancedesk/ms-office-ai-helper.git
cd ms-office-ai-helper

# Install dependencies
npm install

# Install SSL certificates for local development
npx office-addin-dev-certs install

# Start the development server
npm start
```

## 📁 Project Structure

```
ms-office-ai-helper/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html         # Main UI
│   │   ├── taskpane.js           # Main application logic
│   │   └── modules/
│   │       ├── aiContext.js      # System prompts & API reference for AI
│   │       ├── actionExecutor.js # Executes [EXECUTE] & [ACTION] from AI responses
│   │       ├── chatUI.js         # Message rendering, markdown
│   │       ├── settingsPanel.js  # API key setup & provider switch
│   │       └── specialCommands.js # Slash commands
│   ├── commands/                 # Office ribbon commands
│   └── services/
│       ├── groqService.js        # Groq API (Llama 3.3 70B)
│       ├── geminiService.js      # Google Gemini API
│       ├── documentService.js    # Word document operations
│       └── apiKeyManager.js      # API key storage
├── dist/                         # Production build
├── manifest.xml                  # Office Add-in manifest
├── webpack.config.js             # Build configuration
└── package.json
```

## ⚙️ Configuration

### API Keys

1. Click the **Settings ⚙️** button in the add-in
2. Enter your Groq or Gemini API key
3. Click **Test Connection** to verify
4. Click **Save** to store the key securely

API keys are stored locally and are never sent anywhere except the respective AI provider.

### AI Models

- **Groq**: Llama 3.3 70B (most capable production model)
- **Gemini**: Google Gemini Flash

## 💬 Usage Examples

Just type naturally—in any language. The AI interprets your intent and executes the right Word actions:

| What You Say | What Happens |
|--------------|--------------|
| "Write a 2-page essay on MIMD computers" | Generates and inserts content at the end of your document |
| "Find 'important' and make it bold" | Searches and applies bold formatting |
| "Delete everything in the document" | Clears the document |
| "Move the word X to the top" | Finds X, removes it, inserts at start |
| "Underline all instances of 'project'" | Finds and underlines the text |
| "Replace 'old' with 'new'" | Find and replace |
| "Add a heading 'Conclusion' at the end" | Inserts a new heading |
| "Center the first paragraph" | Applies center alignment |
| "Highlight all 'warning' in yellow" | Applies yellow highlight |
| "Summarize this document" | AI summarizes the content |

### Content Writing

When you ask to write an article, essay, or assignment, the AI:

- Inserts content **into your current document** (appends at the end)
- Uses your document's **Normal style** (e.g., Aptos Body 12pt)
- Adds a blank line before new content for clean separation
- Aims for factual, human-sounding text that avoids AI detection

## 🔧 Sideloading the Add-in

### Microsoft 365 / Office 365

1. Open Word
2. Go to **Insert** → **Get Add-ins** → **MY ADD-INS** tab
3. Click **Upload My Add-in**
4. Browse to `manifest.xml` or download from [GitHub Pages](https://lancedesk.github.io/ms-office-ai-helper/manifest.xml)
5. Click **Upload**

### Office 2019 / Office 2016

See the detailed guide: [docs/INSTALLATION.md](docs/INSTALLATION.md)

### Word Online

1. Go to [office.com](https://www.office.com) and open Word Online
2. Open a document
3. **Insert** → **Add-ins** → **Upload Add-in**
4. Upload `manifest.xml`

## 🔧 Troubleshooting

### SSL Certificate Issues (Local Development)

```bash
npx office-addin-dev-certs uninstall
npx office-addin-dev-certs install --days 365
```

### Cache Issues (Add-in Not Updating)

```bash
npm run clear-all-cache
# Then close Word completely and reopen
```

## 🏗️ Building for Production

```bash
npm run build
npm run deploy   # Deploys to GitHub Pages
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [Microsoft Office Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/) - Office.js framework
- [Groq](https://groq.com/) - Fast AI inference (Llama 3.3 70B)
- [Google Gemini](https://ai.google.dev/) - Google's AI model
