# How to Load and Test the Add-in in Microsoft Word

## Prerequisites
- Microsoft Word (2016+, Office 365, or Word Online)
- The development server is running (`npm run dev-server`)

## Method 1: Sideload in Word Desktop (Windows/Mac)

### Windows

1. **Start the Development Server**
   ```bash
   npm run dev-server
   ```
   Server should be running at https://localhost:3001/

2. **Create a Network Share** (or use a local folder)
   - Create a folder for manifests: `C:\AddInManifests`
   - Copy `manifest.xml` to this folder

3. **Add Trusted Catalog in Word**
   - Open Microsoft Word
   - Go to **File > Options > Trust Center > Trust Center Settings**
   - Click **Trusted Add-in Catalogs**
   - In the **Catalog Url** field, enter: `C:\AddInManifests`
   - Click **Add catalog**
   - Check **Show in Menu**
   - Click **OK**

4. **Load the Add-in**
   - Close and reopen Word
   - Go to **Insert > My Add-ins**
   - Click **Shared Folder** tab
   - You should see "AI Helper"
   - Click it to load

### Mac

1. **Start the Development Server**
   ```bash
   npm run dev-server
   ```

2. **Create Manifest Folder**
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   ```

3. **Copy Manifest**
   ```bash
   cp manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```

4. **Load in Word**
   - Open Word
   - Go to **Insert > Add-ins > My Add-ins**
   - Under **DEVELOPER ADD-INS**, you should see "AI Helper"
   - Click to load

## Method 2: Using Office Add-in Debugger (Easiest)

1. **Start with Debugging**
   ```bash
   npm start
   ```
   This will automatically:
   - Start the dev server
   - Register the manifest
   - Open Word with the add-in loaded

2. **View the Add-in**
   - Word will open automatically
   - Look for "AI Helper" button in the Home ribbon
   - Click it to open the task pane

## Method 3: Word Online

1. **Upload Manifest**
   - Go to https://www.office.com/launch/word
   - Create or open a document
   - Click **Insert > Office Add-ins**
   - Click **Upload My Add-in** (at the bottom)
   - Browse and select `manifest.xml`
   - Click **Upload**

2. **Access the Add-in**
   - The add-in should appear in the ribbon
   - Click to open the task pane

## Testing the Add-in

Once loaded, you should see:

1. **Task Pane** with purple gradient header "🤖 AI Helper"
2. **Welcome Screen** showing features
3. **Chat Interface** at the bottom
4. **Input field** to type messages

### Test Commands (Phase 1 - Basic Testing):

- Type **"hello"** - Get a welcome message
- Type **"read document"** - Test document reading (add some text to your doc first)
- Type **"insert text"** - Test inserting text into document
- Type anything else - Get a placeholder response

## Troubleshooting

### SSL Certificate Warning
- When you first load `https://localhost:3001/`, your browser/Word will show a certificate warning
- This is normal for development
- Click **Advanced** and **Proceed** (or trust the certificate)

### Add-in Not Appearing
1. Make sure the dev server is running: `npm run dev-server`
2. Check the URL is correct: https://localhost:3001/taskpane.html
3. Clear Office cache:
   - Windows: Delete contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Mac: Delete `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/`
4. Restart Word

### Errors in Task Pane
1. Open Developer Tools:
   - Right-click in the task pane
   - Select **Inspect** or **Inspect Element**
2. Check the Console for errors
3. Common issues:
   - SSL certificate not trusted
   - Office.js not loading (check internet connection)
   - Port conflict (change port in webpack.config.js and manifest.xml)

### Certificate Issues
Generate trusted certificates:
```bash
npx office-addin-dev-certs install
```

## Next Steps

Once Phase 1 is working:
- You should see the chat interface
- You can type messages
- Basic document operations work
- Ready for Phase 2: Groq API integration

## Development Tips

- **Auto-reload**: Changes to code will auto-reload in Word (Webpack HMR)
- **Console logs**: Use browser dev tools (F12 in task pane)
- **Manifest changes**: Require reloading the add-in
- **Clear cache**: If changes don't appear, clear Office cache and reload

---

**Status**: Phase 1 Complete ✅
**Next**: Phase 2 - Groq API Integration
