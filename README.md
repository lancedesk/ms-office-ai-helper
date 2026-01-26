# MS Office AI Helper

An intelligent AI-powered assistant add-in for Microsoft Word that integrates with Groq and Google Gemini AI to provide document analysis, smart editing, summaries, and formatting capabilities.

## ✨ Features

- 📄 **Document Analysis** - Read and understand your document content
- ✍️ **Smart Editing** - Format, edit, or restructure content with natural language commands
- 💡 **Summaries** - Get quick summaries of your documents
- 🎨 **Formatting** - Create headers, tables, and more with simple commands
- 🔄 **Multi-Provider Support** - Switch between Groq and Google Gemini AI
- 🤖 **Natural Language** - Just tell the AI what you want in plain English

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
4. Enter your API key when prompted

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
│   │   ├── taskpane.html      # Main UI
│   │   └── taskpane.js        # Main application logic
│   ├── commands/              # Office ribbon commands
│   └── services/              # AI services
│       ├── groqService.js     # Groq API integration
│       ├── geminiService.js   # Gemini API integration
│       ├── documentService.js # Word document operations
│       └── apiKeyManager.js   # API key storage
├── dist/                      # Production build
├── manifest.xml               # Office Add-in manifest
├── webpack.config.js          # Build configuration
└── package.json
```

## ⚙️ Configuration

### API Keys

When you first open the add-in, you'll be prompted to enter your API key:

1. Click the **Settings ⚙️** button
2. Enter your Groq or Gemini API key
3. Click **Test Connection** to verify
4. Click **Save** to store the key securely

API keys are stored locally in your browser/Office storage and are never sent to any server except the respective AI provider.

## 💬 Usage Examples

Just type naturally! The AI understands what you want:

| What You Say | What Happens |
|--------------|--------------|
| "Find the word 'important' and make it bold" | Searches and applies bold formatting |
| "Underline all instances of 'COMP 414'" | Finds and underlines the text |
| "Replace 'old text' with 'new text'" | Find and replace |
| "Add a heading called 'Conclusion' at the end" | Inserts a new heading |
| "Highlight all occurrences of 'warning' in yellow" | Applies yellow highlight |
| "Create a new blank document" | Opens a new Word document |
| "Summarize this document" | AI summarizes the content |

## 🔧 Sideloading the Add-in

### Microsoft 365 / Office 365

1. Open Word
2. Go to **Insert** → **Get Add-ins** → **MY ADD-INS** tab
3. Click **Upload My Add-in**
4. Browse to `manifest.xml` or download from [GitHub Pages](https://lancedesk.github.io/ms-office-ai-helper/manifest.xml)
5. Click **Upload**

### Office 2019 / Office 2016

See the detailed guide: [INSTALLATION.md](INSTALLATION.md)

### Word Online

1. Go to Word Online at [office.com](https://www.office.com)
2. Open a document
3. **Insert** → **Add-ins** → **Upload Add-in**
4. Upload `manifest.xml`

## 🔧 Troubleshooting

### SSL Certificate Issues (Local Development)

```bash
# Reinstall certificates
npx office-addin-dev-certs uninstall
npx office-addin-dev-certs install --days 365
```

### Cache Issues

```bash
# Windows - clear Office add-in caches
powershell -ExecutionPolicy Bypass -File scripts/clear-all-caches.ps1

# Then restart
npm run start-clean
```

## 🏗️ Building for Production

```bash
npm run build
npm run deploy  # Deploys to GitHub Pages
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
- [Groq](https://groq.com/) - Fast AI inference
- [Google Gemini](https://ai.google.dev/) - Google's AI model
