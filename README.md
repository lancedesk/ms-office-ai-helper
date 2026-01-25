# MS Office AI Helper

An intelligent AI-powered assistant add-in for Microsoft Word that integrates with Groq and Google Gemini AI to provide document analysis, smart editing, summaries, and formatting capabilities.

## ✨ Features

- 📄 **Document Analysis** - Read and understand your document content
- ✍️ **Smart Editing** - Format, edit, or restructure content with AI assistance
- 💡 **Summaries** - Get quick summaries of your documents
- 🎨 **Formatting** - Create headers, tables, and more with simple commands
- 🔄 **Multi-Provider Support** - Switch between Groq and Google Gemini AI

## 📋 Prerequisites

- Microsoft Word 2016 or later (Windows/Mac) or Word Online
- Node.js 18+ and npm
- A Groq API key ([Get one here](https://console.groq.com/keys)) or Google Gemini API key ([Get one here](https://aistudio.google.com/apikey))

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/lancedesk/ms-office-ai-helper.git
cd ms-office-ai-helper

# Install dependencies
npm install

# Install SSL certificates for local development
npx office-addin-dev-certs install
```

### Development

```bash
# Start the development server and sideload in Word
npm start

# Or run just the dev server
npm run dev-server
```

### Production Build

```bash
npm run build
```

## ⚙️ Configuration

### API Keys

When you first open the add-in, you'll be prompted to enter your API key:

1. Click the **Settings ⚙️** button
2. Enter your Groq or Gemini API key
3. Click **Test Connection** to verify
4. Click **Save** to store the key securely

API keys are stored locally in your browser/Office storage and are never sent to any server except the respective AI provider.

## 📁 Project Structure

```
ms-office-ai-helper/
├── src/
│   ├── taskpane/          # Main UI (taskpane.html, taskpane.js)
│   ├── commands/          # Office commands
│   └── services/          # AI services (Groq, Gemini, API Key Manager)
├── scripts/               # Utility scripts
├── docs/                  # Documentation
├── manifest.xml           # Office Add-in manifest
├── webpack.config.js      # Build configuration
└── package.json
```

## 🌐 Deployment

### GitHub Pages (Recommended)

This add-in is deployed to GitHub Pages for reliable HTTPS hosting:

**Live URL:** `https://lancedesk.github.io/ms-office-ai-helper/`

To deploy your own:

```bash
npm run build
npm run deploy
```

Then update the `manifest.xml` URLs to point to your GitHub Pages URL.

### Manual Deployment

1. Run `npm run build`
2. Upload the `dist/` folder contents to any HTTPS-enabled web server
3. Update all URLs in `manifest.xml` to point to your server
4. Sideload the updated manifest in Word

## 🔧 Troubleshooting

### SSL Certificate Issues (Local Development)

If you see certificate warnings:

```bash
# Reinstall certificates
npx office-addin-dev-certs uninstall
npx office-addin-dev-certs install --days 365

# Run the loopback setup (Windows, as Administrator)
powershell -ExecutionPolicy Bypass -File scripts/setup-loopback.ps1
```

### Cache Issues

Clear Office add-in caches:

```bash
# Windows
powershell -ExecutionPolicy Bypass -File scripts/clear-all-caches.ps1

# Then restart
npm run start-clean
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
