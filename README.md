# Enhanced Docx/ODT Viewer for VS Code

![Docx Viewer](assets/icons/icon.png)

![Installs](https://img.shields.io/visual-studio-marketplace/i/ShahilKumar.docxreader)
![Visual Studio Marketplace Version](https://img.shields.io/visual-studio-marketplace/v/ShahilKumar.docxreader.svg)
![Rating](https://img.shields.io/visual-studio-marketplace/r/ShahilKumar.docxreader.svg)
![LICENSE](https://img.shields.io/github/license/skfrost19/Docx-Viewer.svg)
[![GitHub Release](https://img.shields.io/github/release/skfrost19/Docx-Viewer.svg)](https://github.com/skfrost19/Docx-Viewer/releases)
![Open Issues](https://img.shields.io/github/issues-raw/skfrost19/Docx-Viewer.svg)
![Closed Issues](https://img.shields.io/github/issues-closed-raw/skfrost19/Docx-Viewer.svg)

A modern, feature-rich document viewer for Microsoft Word (.docx) and OpenDocument Text (.odt) files in Visual Studio Code.

> **About this fork.** This is a fork of [Docx/ODT Viewer](https://github.com/skfrost19/Docx-Viewer) by **Shahil Kumar** ([@skfrost19](https://github.com/skfrost19)), released under the MIT License. All credit for the original extension goes to him; the badges above refer to the original project. The changes made in this fork are listed below.

## 🆕 What's new in this fork

Changes relative to the original extension (v1.4.3):

### Word-accurate layout
- **Real pages** - the page width, height and margins come from the document itself; the text is split into pages as in Word (a paragraph that does not fit continues on the next page), Word's own and manual page breaks are honoured, and pages are numbered
- **Paragraph formatting from Word styles** - alignment, spacing before/after, line spacing, indents, font size, bold and italic are resolved per paragraph from direct formatting, the style chain and the document defaults; the document's own font is used
- **Character formatting** - text colour, highlight, shading, size, font, underline, small/all caps, hidden text and character-style bold/italic
- **Word spacing model** - space before/after added the way Word does (no CSS margin collapse), "don't add space between paragraphs of the same style" respected, and line height taken from the font itself instead of a fixed 1.15

### Equations
- **Word equations (OMML)** are converted to MathML and rendered natively, both inline and as display equations

### Comments and tracked changes
- **Comments** shown as cards beside the text, with replies and resolved state; the commented text is highlighted
- **Tracked changes** shown as insertions and deletions, including nested changes and images inside a change
- **Three views**, switched with one button: *Markup* (changes and comments), *Final* (final text with comments) and *Clean* (final text only)

### Live reload
- **The document re-renders when the file changes on disk** (saved from Word, synced by OneDrive/iCloud); the scroll position and the chosen view are kept. Because cloud folders do not always report changes, the file size and time are also checked every 2 s

### Look and controls
- **Light/dark toggle** - light mode by default; one click switches between light and dark, and the choice is kept across reloads (`docxreader.theme` still accepts `auto` to follow VS Code)
- **Redesigned toolbar** - a single floating bar with line icons instead of emoji, grouped as outline | zoom | find, view, theme | collapse, shown by default; it collapses into a small pill with the theme switch
- **Find panel** restyled to match the toolbar

### Install this fork
This fork is not on the VS Code Marketplace. Build and install it from source:

```bash
git clone https://github.com/kanak-buet19/docx-viewer.git
cd docx-viewer
npm install && npm run compile
npx @vscode/vsce package
code --install-extension docxreader-<version>.vsix
```

## ✨ Features

### 📄 Document Support
- **Microsoft Word Documents (.docx)** - Full support with proper formatting
- **OpenDocument Text (.odt)** - Complete compatibility with LibreOffice/OpenOffice documents

### 🎨 Modern UI & Experience
- **VS Code Theme Integration** - Automatically adapts to your VS Code theme (dark/light/high contrast)
- **Toolbar Toggle** - Hide/show toolbar for distraction-free reading experience
- **Document Outline** - Navigate through headings with an interactive sidebar
- **Search Functionality** - Find text within documents with highlighting
- **Zoom Controls** - Zoom in/out with keyboard shortcuts or toolbar buttons

### 🔧 Enhanced Functionality
- **Responsive Design** - Adapts to different panel sizes
- **Keyboard Shortcuts** - Full keyboard navigation support
- **Status Bar Integration** - Shows current zoom level
- **Error Handling** - Graceful error messages and loading states
- **Performance Optimized** - Fast rendering with modern CSS techniques

### ⌨️ Keyboard Shortcuts
- `Ctrl+Plus` / `Cmd+Plus` - Zoom In
- `Ctrl+Minus` / `Cmd+Minus` - Zoom Out  
- `Ctrl+0` / `Cmd+0` - Reset Zoom
- `Ctrl+F` / `Cmd+F` - Search in Document
- `Ctrl+H` / `Cmd+H` - Toggle Toolbar
- `Escape` - Close Search Panel

### 🛠️ Toolbar Features
- **Zoom Controls** - Precise zoom adjustment (50% to 300%)
- **Outline Toggle** - Show/hide document structure
- **Theme Toggle** - Cycle through Auto (🔄) / Light (🌙) / Dark (☀️) modes
- **Search** - Quick text search with navigation
- **Toolbar Toggle** - Hide toolbar for distraction-free reading

## 🚀 Installation

1. Install from the VS Code Marketplace
2. Open any `.docx` or `.odt` file
3. The document will automatically open in the enhanced viewer

## ⚙️ Configuration

Access settings via `Ctrl+Shift+P` → "Open Docx Reader Configuration" or go to File → Preferences → Settings and search for "Docx Reader".

### Available Settings

| Setting | Description | Default |
|---------|-------------|---------|
| `docxreader.font` | Font family for document rendering | `Arial` |
| `docxreader.theme` | Starting theme (light/dark/auto). Auto follows the VS Code theme. The ☀️/🌙 button in the viewer switches between light and dark | `light` |
| `docxreader.zoomLevel` | Default zoom level (0.5 - 3.0) | `1.0` |
| `docxreader.showOutline` | Show document outline by default | `true` |

### Example Configuration
```json
{
    "docxreader.font": "Georgia",
    "docxreader.theme": "light",
    "docxreader.zoomLevel": 1.2,
    "docxreader.showOutline": true
}
```

## 🎯 Usage

### Opening Documents
1. **Right-click** any `.docx` or `.odt` file → "Open With" → "Docx Reader"
2. **Double-click** the file (if set as default viewer)
3. Use **Command Palette** (`Ctrl+Shift+P`) → "View: Reopen Editor With" → "Docx Reader"

### Navigation
- **Outline Panel**: Click headings to jump to sections
- **Search**: Use toolbar search or `Ctrl+F` to find text
- **Zoom**: Use toolbar buttons or keyboard shortcuts
- **Toolbar Toggle**: Click the ✕ button to hide toolbar, or press `Ctrl+H` to toggle

### Distraction-Free Reading
For a cleaner reading experience:
1. **Hide Toolbar**: Click the ✕ button or press `Ctrl+H`
2. **Show Toolbar**: Click the ⚙️ gear icon that appears in the top-right corner
3. **Quick Toggle**: Use `Ctrl+H` keyboard shortcut anytime

### Document Structure
The viewer automatically extracts and displays:
- Headings (H1-H6) in the outline
- Formatted text with proper styling  
- Tables with borders and styling
- Images with responsive sizing
- Lists (numbered and bulleted)
- Block quotes and code blocks

## 🛠️ Development

### Building from Source
```bash
# Clone repository
git clone https://github.com/skfrost19/Docx-Viewer.git
cd Docx-Viewer

# Install dependencies  
npm install

# Compile TypeScript
npm run compile

# Run tests
npm test

# Launch development version
code --extensionDevelopmentPath=. --new-window
```

### Technology Stack
- **TypeScript 5.8+** - Modern JavaScript with strong typing
- **VS Code API 1.90+** - Latest VS Code extension capabilities
- **Mammoth.js** - DOCX to HTML conversion
- **ODT2HTML** - ODT file processing
- **Modern CSS** - Responsive design with CSS Grid/Flexbox

## 🐛 Troubleshooting

### Common Issues

**Document not loading:**
- Ensure the file isn't corrupted
- Try opening with another application first
- Check file permissions

**Outline navigation issues:**
- If clicking on outline items doesn't navigate correctly, ensure you're using version 1.2.2 or later
- For documents with duplicate heading names, the extension now correctly navigates to the specific section (fixed in v1.2.2)

**Toolbar visibility issues:**
- If toolbar doesn't hide/show properly, ensure you're using version 1.3.0 or later
- The mini toggle button (⚙️) appears in the top-right corner when toolbar is hidden
- Use `Ctrl+H` keyboard shortcut as an alternative to button clicks

**Styling issues:**
- Verify VS Code theme compatibility
- Reset zoom to 100%
- Check font settings in configuration

**Performance problems:**
- Close other heavy extensions
- Reduce zoom level for large documents
- Enable hardware acceleration in VS Code

### Getting Help
1. Check existing [Issues](https://github.com/skfrost19/Docx-Viewer/issues)
2. Create a new issue with:
   - VS Code version
   - Extension version
   - Sample document (if possible)
   - Error messages

## 📝 Changelog
See [Changelog](https://github.com/skfrost19/Docx-Viewer/blob/main/CHANGELOG.md)

## 🤝 Contributing

Contributions are welcome! Please read our [Contributing Guidelines](CONTRIBUTING.md) first.

### Areas for Contribution
- Additional file format support
- UI/UX improvements  
- Performance optimizations
- Documentation improvements
- Bug fixes and testing

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE.txt) file for details.

## 🙏 Acknowledgments

- [Mammoth.js](https://github.com/mwilliamson/mammoth.js/) for DOCX conversion
- [ODT2HTML](https://github.com/odt2html/odt2html) for ODT processing
- VS Code team for the excellent extension API
- Contributors and users for feedback and suggestions

---

**Made with ❤️ for the VS Code community**
