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

## � Sideloading the Add-in

### Microsoft 365 / Office 365 (Recommended)

1. Open Word
2. Go to **Insert** → **Get Add-ins** → **MY ADD-INS** tab
3. Click **Upload My Add-in**
4. Browse to `manifest.xml` (or download from GitHub Pages)
5. Click **Upload**

### Office 2019 / Office 2016 (Shared Folder Method)

Office 2019/2016 requires a **Trusted Catalog** instead of direct upload:

#### Step 1: Create a Shared Folder Catalog

1. Create a folder on your computer, e.g., `C:\OfficeAddins`
2. Share the folder:
   - Right-click → **Properties** → **Sharing** tab
   - Click **Share...** → Add **Everyone** with **Read** permissions
   - Note the network path (e.g., `\\YOUR-PC-NAME\OfficeAddins`)

#### Step 2: Add as Trusted Catalog

1. Open Word → **File** → **Options** → **Trust Center**
2. Click **Trust Center Settings...** → **Trusted Add-in Catalogs**
3. In **Catalog Url**, enter your shared folder path: `\\YOUR-PC-NAME\OfficeAddins`
4. Click **Add catalog** → Check **Show in Menu** checkbox
5. Click **OK** → **OK** again
6. **Restart Word**

#### Step 3: Install the Add-in

1. Download `manifest.xml` from `https://lancedesk.github.io/ms-office-ai-helper/manifest.xml`
2. Copy `manifest.xml` to your shared folder (`C:\OfficeAddins`)
3. Open Word → **Insert** → **My Add-ins** → **SHARED FOLDER** tab
4. Select **AI Helper** → Click **Add**

### Word Online

1. Go to Word Online at [office.com](https://www.office.com)
2. Open a document
3. **Insert** → **Add-ins** → **Upload Add-in**
4. Upload `manifest.xml`

## �🔧 Troubleshooting

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
