import * as vscode from 'vscode';
import * as path from 'path';
import { DocxHandler } from './docx_handler';
import { OdtHandler } from './odt_handler';

export class DocumentRenderer {
    private static currentZoom: number = 1.0;
    private static outlineVisible: boolean = true;
    private static currentTheme: string = 'auto';
    private static toolbarVisible: boolean = false;

    public static async renderDocument(uri: vscode.Uri, panel: vscode.WebviewPanel, initialState?: { zoom?: number; outlineVisible?: boolean; toolbarVisible?: boolean }) {
        const docxPath = uri.fsPath;
        try {
            // Show loading state
            panel.webview.html = this.getLoadingHtml();

            // Get configuration
            // Config theme setting takes precedence over system theme (Fixes #27, refixes #25)
            // - 'auto': follow VS Code system theme (auto-detect)
            // - 'light': always use light theme (ignore system theme)
            // - 'dark': always use dark theme (ignore system theme)
            const config = vscode.workspace.getConfiguration('docxreader');
            const font = config.get('font', 'Arial');
            const theme = config.get('theme', 'auto');
            // Use caller-supplied state when provided, otherwise fall back to config defaults.
            // NOTE: theme always comes from config (config takes precedence over runtime toggle state)
            // to preserve the fix for #27 / #25.
            this.currentZoom = initialState?.zoom ?? config.get('zoomLevel', 1.0);
            this.outlineVisible = initialState?.outlineVisible ?? config.get('showOutline', true);
            this.currentTheme = theme; // always config — never initialState.theme
            if (initialState?.toolbarVisible !== undefined) {
                this.toolbarVisible = initialState.toolbarVisible;
            }

            let documentHtml = '';
            let processedData: { html: string; outline: OutlineItem[] };

            if (docxPath.toLowerCase().endsWith('.docx')) {
                documentHtml = await DocxHandler.renderDocx(uri);
                processedData = this.processDocumentHtmlAndExtractOutline(documentHtml);
            } else if (docxPath.toLowerCase().endsWith('.odt')) {
                documentHtml = await OdtHandler.renderOdt(uri);
                processedData = this.processDocumentHtmlAndExtractOutline(documentHtml);
            } else {
                panel.webview.html = this.getErrorHtml('Unsupported file format', 'Only .docx and .odt files are supported.');
                return;
            }

            // Generate enhanced HTML with modern UI
            panel.webview.html = this.generateEnhancedHtml(processedData.html, processedData.outline, font, theme, docxPath);

        } catch (error) {
            console.error('Error rendering document:', error);
            panel.webview.html = this.getErrorHtml('Failed to load document', error instanceof Error ? error.message : String(error));
        }
    }

    private static generateEnhancedHtml(documentHtml: string, outline: OutlineItem[], font: string, theme: string, filePath: string): string {
        const fileName = path.basename(filePath);

        // Determine theme class (still set on body for auto-detection CSS fallback)
        let themeClass = '';
        if (this.currentTheme === 'dark') {
            themeClass = 'vscode-dark';
        } else if (this.currentTheme === 'light') {
            themeClass = 'vscode-light';
        } else {
            // 'auto' - detect VS Code theme based on CSS variables
            themeClass = 'vscode-theme-auto';
        }

        const outlineHtml = this.generateOutlineHtml(outline);

        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${fileName} - Docx Viewer</title>
    <style>
        ${this.getInlineCSS()}
    </style>
    <!-- Theme override: injected by config. Uses !important so VS Code body-class injections cannot override it.
         For 'auto', this block is empty and body-class-based detection takes over. (Fixes #27, refixes #25) -->
    <style id="docx-theme-override">${this.getThemeOverrideCSS(this.currentTheme)}</style>
</head>
<body class="${themeClass}" style="font-family: ${font};">
    <div class="docx-viewer-container">
        <!-- Toolbar -->
        <div class="docx-toolbar ${this.toolbarVisible ? '' : 'hidden'}" id="mainToolbar">
            <button id="zoomOut" title="Zoom Out">🔍-</button>
            <span id="zoomLevel">${Math.round(this.currentZoom * 100)}%</span>
            <button id="zoomIn" title="Zoom In">🔍+</button>
            <button id="resetZoom" title="Reset Zoom">⚊</button>
            <button id="toggleOutline" title="Toggle Outline">${this.outlineVisible ? '◧' : '◨'}</button>
            <button id="themeToggle" title="Toggle Theme">${this.currentTheme === 'dark' ? '☀️' : this.currentTheme === 'light' ? '🌙' : '🔄'}</button>
            <button id="searchBtn" title="Search">🔍</button>
            <button id="toggleMarkup" title="Markup: tracked changes and comments (click: final)">±</button>
            <button id="hideToolbar" title="Hide Toolbar">✕</button>
        </div>
        
        <!-- Mini Toolbar Toggle (visible when main toolbar is hidden) -->
        <div class="docx-mini-toolbar ${this.toolbarVisible ? 'hidden' : ''}" id="miniToolbar">
            <button id="showToolbar" title="Show Toolbar">⚙️</button>
        </div>
        
        <!-- Search Panel -->
        <div class="docx-search" id="searchPanel">
            <input type="text" id="searchInput" placeholder="Search in document...">
            <span id="searchCount" class="search-count">0/0</span>
            <button id="searchPrev" title="Previous">↑</button>
            <button id="searchNext" title="Next">↓</button>
            <button id="closeSearch" title="Close">✕</button>
        </div>
        
        <!-- Outline Panel -->
        <div class="docx-outline ${this.outlineVisible ? '' : 'hidden'}" id="outlinePanel">
            <h3>📋 Document Outline</h3>
            ${outlineHtml}
        </div>
        
        <!-- Document Content -->
        <div class="docx-content" id="documentContent">
            <div class="docx-document" id="document" style="transform: scale(${this.currentZoom});">
                ${documentHtml}
            </div>
        </div>
    </div>
    
    <script>
        ${this.getViewerScript()}
    </script>
</body>
</html>`;
    }

    /**
     * Returns CSS that hardcodes the viewer theme variables directly on :root using !important.
     * This takes precedence over any body class VS Code injects into the webview (e.g. vscode-dark).
     * Returns empty string for 'auto' — in that case body-class detection continues to work. (Fixes #27)
     */
    private static getThemeOverrideCSS(theme: string): string {
        if (theme === 'light') {
            return `
        :root {
            --viewer-bg: #ffffff !important;
            --viewer-fg: #000000 !important;
            --viewer-border: #cccccc !important;
            --viewer-hover: #f0f0f0 !important;
            --viewer-active: #0078d4 !important;
            --viewer-shadow: rgba(0, 0, 0, 0.1) !important;
            --document-bg: #ffffff !important;
        }
        body { background: #ffffff !important; color: #000000 !important; }
        .docx-document { background: #ffffff !important; color: #000000 !important; }
        `;
        }
        if (theme === 'dark') {
            return `
        :root {
            --viewer-bg: #1e1e1e !important;
            --viewer-fg: #ffffff !important;
            --viewer-border: #3c3c3c !important;
            --viewer-hover: #2a2d2e !important;
            --viewer-active: #0078d4 !important;
            --viewer-shadow: rgba(255, 255, 255, 0.1) !important;
            --document-bg: #2d2d30 !important;
        }
        body { background: #1e1e1e !important; color: #ffffff !important; }
        .docx-document { background: #2d2d30 !important; color: #ffffff !important; }
        `;
        }
        // 'auto': no override — body-class detection handles it
        return '';
    }

    private static processDocumentHtmlAndExtractOutline(html: string): { html: string; outline: OutlineItem[] } {
        const outline: OutlineItem[] = [];
        let headingCounter = 0;

        // Process HTML and extract outline in a single pass to ensure matching IDs
        const processedHtml = html.replace(/<(h[1-6])[^>]*>(.*?)<\/h[1-6]>/gi, (match, tag, content) => {
            const level = parseInt(tag.charAt(1));
            const text = content.replace(/<[^>]*>/g, '').trim();
            const id = this.generateHeadingId(text, headingCounter++);

            if (text) {
                outline.push({
                    level,
                    text,
                    id
                });
            }

            return `<${tag} id="${id}">${content}</${tag}>`;
        });

        return { html: processedHtml, outline };
    }

    private static generateHeadingId(text: string, index: number): string {
        const baseId = text.toLowerCase()
            .replace(/[^\w\s-]/g, '')
            .replace(/\s+/g, '-')
            .substring(0, 50);

        // Make ID unique by adding index
        return `${baseId}-${index}`;
    }

    private static generateOutlineHtml(outline: OutlineItem[]): string {
        if (outline.length === 0) {
            return '<p style="opacity: 0.7; font-style: italic;">No headings found</p>';
        }

        return outline.map(item =>
            `<div class="docx-outline-item level-${item.level}" data-target="${item.id}">
                ${item.text}
            </div>`
        ).join('');
    }

    private static getLoadingHtml(): string {
        return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            margin: 0; 
            padding: 0; 
            font-family: var(--vscode-font-family); 
            background: var(--vscode-editor-background);
            color: var(--vscode-editor-foreground);
        }
        .docx-loading {
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            font-size: 14px;
        }
        .spinner {
            width: 20px;
            height: 20px;
            border: 2px solid var(--vscode-input-border);
            border-top: 2px solid var(--vscode-button-background);
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin-left: 10px;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    </style>
</head>
<body>
    <div class="docx-loading">
        Loading document...
        <div class="spinner"></div>
    </div>
</body>
</html>`;
    }

    private static getErrorHtml(title: string, message: string): string {
        return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            margin: 0; 
            padding: 20px; 
            font-family: var(--vscode-font-family); 
            background: var(--vscode-editor-background);
            color: var(--vscode-editor-foreground);
        }
        .docx-error {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            height: 80vh;
            text-align: center;
        }
        .error-icon { font-size: 48px; margin-bottom: 20px; }
        h2 { color: var(--vscode-errorForeground); margin: 0 0 10px 0; }
        p { opacity: 0.8; margin: 0; }
    </style>
</head>
<body>
    <div class="docx-error">
        <div class="error-icon">⚠️</div>
        <h2>${title}</h2>
        <p>${message}</p>
    </div>
</body>
</html>`;
    }

    private static getInlineCSS(): string {
        return `
        /* Embedded CSS for the document viewer */
        :root {
            --viewer-bg: var(--vscode-editor-background, #ffffff);
            --viewer-fg: var(--vscode-editor-foreground, #000000);
            --viewer-border: var(--vscode-panel-border, #cccccc);
            --viewer-hover: var(--vscode-list-hoverBackground, #f0f0f0);
            --viewer-active: var(--vscode-list-activeSelectionBackground, #0078d4);
            --viewer-shadow: rgba(0, 0, 0, 0.1);
            --outline-width: 250px;
            --document-bg: #ffffff;
        }
        
        body.vscode-light {
            --viewer-bg: #ffffff;
            --viewer-fg: #000000;
            --viewer-border: #cccccc;
            --viewer-hover: #f0f0f0;
            --viewer-active: #0078d4;
            --viewer-shadow: rgba(0, 0, 0, 0.1);
            --document-bg: #ffffff;
        }
        
        body.vscode-dark {
            --viewer-bg: #1e1e1e;
            --viewer-fg: #ffffff;
            --viewer-border: #3c3c3c;
            --viewer-hover: #2a2d2e;
            --viewer-active: #0078d4;
            --viewer-shadow: rgba(255, 255, 255, 0.1);
            --document-bg: #2d2d30;
        }
        
        * { box-sizing: border-box; }
        body { margin: 0; padding: 0; overflow: hidden; }
        
        .docx-viewer-container {
            display: flex;
            height: 100vh;
            background: var(--viewer-bg);
            color: var(--viewer-fg);
        }
        
        .docx-toolbar {
            position: fixed;
            top: 0;
            right: 0;
            z-index: 1000;
            display: flex;
            gap: 8px;
            padding: 8px;
            background: var(--viewer-bg);
            border-bottom: 1px solid var(--viewer-border);
            border-left: 1px solid var(--viewer-border);
            border-radius: 0 0 0 4px;
            box-shadow: 0 2px 8px var(--viewer-shadow);
        }
        
        .docx-toolbar button {
            background: transparent;
            border: 1px solid var(--viewer-border);
            color: var(--viewer-fg);
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
            transition: all 0.2s ease;
        }
        
        .docx-toolbar button:hover { background: var(--viewer-hover); }
        .docx-toolbar button:active { background: var(--viewer-active); color: white; }
        
        .docx-toolbar.hidden {
            transform: translateY(-100%);
            opacity: 0;
            pointer-events: none;
        }
        
        .docx-mini-toolbar {
            position: fixed;
            top: 8px;
            right: 8px;
            z-index: 1001;
            background: var(--viewer-bg);
            border: 1px solid var(--viewer-border);
            border-radius: 4px;
            padding: 4px;
            box-shadow: 0 2px 8px var(--viewer-shadow);
            transition: all 0.3s ease;
        }
        
        .docx-mini-toolbar.hidden {
            opacity: 0;
            pointer-events: none;
            transform: scale(0.8);
        }
        
        .docx-mini-toolbar button {
            background: transparent;
            border: 1px solid var(--viewer-border);
            color: var(--viewer-fg);
            padding: 8px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.2s ease;
        }
        
        .docx-mini-toolbar button:hover { background: var(--viewer-hover); }
        .docx-mini-toolbar button:active { background: var(--viewer-active); color: white; }
        
        #zoomLevel {
            padding: 6px 8px;
            font-size: 12px;
            min-width: 45px;
            text-align: center;
        }
        
        .docx-outline {
            width: var(--outline-width);
            background: var(--viewer-bg);
            border-right: 1px solid var(--viewer-border);
            overflow-y: auto;
            padding: 16px;
            transition: margin-left 0.3s ease;
        }
        
        .docx-outline.hidden { margin-left: calc(-1 * var(--outline-width)); }
        
        .docx-outline h3 {
            margin: 0 0 12px 0;
            font-size: 14px;
            font-weight: 600;
            border-bottom: 1px solid var(--viewer-border);
            padding-bottom: 8px;
        }
        
        .docx-outline-item {
            padding: 4px 8px;
            cursor: pointer;
            border-radius: 4px;
            font-size: 12px;
            margin: 2px 0;
            transition: background 0.2s ease;
        }
        
        .docx-outline-item:hover { background: var(--viewer-hover); }
        .docx-outline-item.active { background: var(--viewer-active); color: white; }
        
        .docx-outline-item.level-1 { padding-left: 8px; font-weight: 600; }
        .docx-outline-item.level-2 { padding-left: 20px; }
        .docx-outline-item.level-3 { padding-left: 32px; }
        .docx-outline-item.level-4 { padding-left: 44px; }
        .docx-outline-item.level-5 { padding-left: 56px; }
        .docx-outline-item.level-6 { padding-left: 68px; }
        
        .docx-content {
            flex: 1;
            overflow: auto;
            padding: 16px;
            scroll-behavior: smooth;
        }
        
        .docx-document {
            max-width: 210mm;
            margin: 0 auto;
            background: var(--document-bg);
            color: var(--viewer-fg);
            box-shadow: 0 0 20px var(--viewer-shadow);
            border-radius: 4px;
            padding: 40px;
            line-height: 1.6;
            transform-origin: top center;
            transition: transform 0.2s ease;
        }
        
        .docx-document h1, .docx-document h2, .docx-document h3,
        .docx-document h4, .docx-document h5, .docx-document h6 {
            margin: 1.5em 0 0.5em 0;
            line-height: 1.4;
        }
        
        .docx-document p { margin: 0.8em 0; }
        .docx-document img { max-width: 100%; height: auto; margin: 1em 0; }
        
        .docx-document table {
            border-collapse: collapse;
            margin: 1em 0;
            width: 100%;
        }
        
        .docx-document table td, .docx-document table th {
            border: 1px solid var(--viewer-border);
            padding: 8px 12px;
        }
        
        .docx-document table th {
            background: var(--viewer-hover);
            font-weight: 600;
        }
        
        .zoom-50 { transform: scale(0.5); }
        .zoom-75 { transform: scale(0.75); }
        .zoom-90 { transform: scale(0.9); }
        .zoom-100 { transform: scale(1.0); }
        .zoom-110 { transform: scale(1.1); }
        .zoom-125 { transform: scale(1.25); }
        .zoom-150 { transform: scale(1.5); }
        .zoom-200 { transform: scale(2.0); }
        .zoom-300 { transform: scale(3.0); }
        
        .docx-search {
            position: fixed;
            top: 50px;
            right: 10px;
            z-index: 999;
            display: none;
            background: var(--viewer-bg);
            border: 1px solid var(--viewer-border);
            border-radius: 4px;
            padding: 8px;
            box-shadow: 0 4px 12px var(--viewer-shadow);
            gap: 4px;
        }
        
        .docx-search.visible { display: flex; }
        
        .docx-search input {
            background: var(--viewer-bg);
            border: 1px solid var(--viewer-border);
            color: var(--viewer-fg);
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            width: 200px;
        }

        .search-count {
            min-width: 42px;
            text-align: center;
            font-size: 12px;
            font-variant-numeric: tabular-nums;
            color: var(--viewer-fg);
            opacity: 0.8;
            padding: 4px 2px;
        }
        
        .docx-search button {
            background: transparent;
            border: 1px solid var(--viewer-border);
            color: var(--viewer-fg);
            padding: 4px 8px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        }
        
        .docx-search button:hover { background: var(--viewer-hover); }
        
        .search-highlight {
            background: yellow;
            color: black;
            padding: 1px 2px;
            border-radius: 2px;
        }
        
        body.vscode-dark .search-highlight {
            background: #ff6600;
            color: white;
        }

        .search-highlight.current {
            background: #ff3b30;
            color: white;
        }

        .diff-added {
            background-color: #e6ffec;
            border-left: 4px solid #2ea043;
        }
        body.vscode-dark .diff-added {
            background-color: rgba(46, 160, 67, 0.15);
            border-left: 4px solid #2ea043;
        }

        .diff-removed {
            background-color: #ffebe9;
            border-left: 4px solid #ff0000;
        }
        body.vscode-dark .diff-removed {
            background-color: rgba(255, 0, 0, 0.15);
            border-left: 4px solid #ff0000;
        }
        `;
    }

    private static getViewerScript(): string {
        return `
        (function() {
            let currentZoom = ${this.currentZoom};
            let outlineVisible = ${this.outlineVisible};
            let currentTheme = '${this.currentTheme}';
            let toolbarVisible = ${this.toolbarVisible};
            let searchResults = [];
            let currentSearchIndex = -1;
            let activeSearchElement = null;

            // Get VS Code API for messaging
            const vscode = acquireVsCodeApi();
            
            // Auto-detect VS Code theme ONLY when theme is set to 'auto' in config
            // When theme is explicitly set to 'light' or 'dark', this returns null
            // to ensure config setting takes precedence over system theme (Fixes #27, refixes #25)
            function detectVSCodeTheme() {
                if (currentTheme === 'auto') {
                    // Get computed style from body to read VS Code CSS variables
                    const computedStyle = getComputedStyle(document.body);
                    const bgColor = computedStyle.getPropertyValue('--vscode-editor-background').trim();
                    
                    // If we can't detect, try to determine from the background color
                    if (!bgColor) {
                        // Fallback: check body background color
                        const bodyBg = computedStyle.backgroundColor;
                        const rgb = bodyBg.match(/\d+/g);
                        if (rgb && rgb.length >= 3) {
                            const brightness = (parseInt(rgb[0]) * 299 + parseInt(rgb[1]) * 587 + parseInt(rgb[2]) * 114) / 1000;
                            return brightness > 128 ? 'vscode-light' : 'vscode-dark';
                        }
                    }
                    
                    // Parse the background color to determine if it's dark or light
                    if (bgColor.startsWith('#')) {
                        // Hex color
                        const hex = bgColor.substring(1);
                        const r = parseInt(hex.substring(0, 2), 16);
                        const g = parseInt(hex.substring(2, 4), 16);
                        const b = parseInt(hex.substring(4, 6), 16);
                        const brightness = (r * 299 + g * 587 + b * 114) / 1000;
                        return brightness > 128 ? 'vscode-light' : 'vscode-dark';
                    } else if (bgColor.startsWith('rgb')) {
                        // RGB color
                        const rgb = bgColor.match(/\d+/g);
                        if (rgb && rgb.length >= 3) {
                            const brightness = (parseInt(rgb[0]) * 299 + parseInt(rgb[1]) * 587 + parseInt(rgb[2]) * 114) / 1000;
                            return brightness > 128 ? 'vscode-light' : 'vscode-dark';
                        }
                    }
                    
                    // Default to light if we can't determine
                    return 'vscode-light';
                }
                return null;
            }
            
            // Apply the config-defined theme to the body, removing any conflicting classes.
            // This is the single source of truth for theme enforcement.
            // VS Code's webview runtime automatically injects body classes (e.g. 'vscode-dark')
            // AFTER the HTML loads, which would override our set class. We must re-enforce
            // the config theme by stripping VS Code's injected classes. (Fixes #27, refixes #25)
            function applyConfigTheme() {
                const body = document.body;

                // Disconnect observer before touching the DOM to avoid triggering ourselves.
                // MutationObserver fires asynchronously so a simple synchronous flag is not enough.
                observer.disconnect();

                // Primary fix: update the :root override <style> block.
                // !important rules here beat any body class VS Code injects. (Fixes #27, refixes #25)
                const overrideStyle = document.getElementById('docx-theme-override');
                if (overrideStyle) {
                    if (currentTheme === 'light') {
                        overrideStyle.textContent = \`
                            :root {
                                --viewer-bg: #ffffff !important;
                                --viewer-fg: #000000 !important;
                                --viewer-border: #cccccc !important;
                                --viewer-hover: #f0f0f0 !important;
                                --viewer-active: #0078d4 !important;
                                --viewer-shadow: rgba(0, 0, 0, 0.1) !important;
                                --document-bg: #ffffff !important;
                            }
                            body { background: #ffffff !important; color: #000000 !important; }
                            .docx-document { background: #ffffff !important; color: #000000 !important; }
                        \`;
                    } else if (currentTheme === 'dark') {
                        overrideStyle.textContent = \`
                            :root {
                                --viewer-bg: #1e1e1e !important;
                                --viewer-fg: #ffffff !important;
                                --viewer-border: #3c3c3c !important;
                                --viewer-hover: #2a2d2e !important;
                                --viewer-active: #0078d4 !important;
                                --viewer-shadow: rgba(255, 255, 255, 0.1) !important;
                                --document-bg: #2d2d30 !important;
                            }
                            body { background: #1e1e1e !important; color: #ffffff !important; }
                            .docx-document { background: #2d2d30 !important; color: #ffffff !important; }
                        \`;
                    } else {
                        // 'auto': clear override, let body-class detection work
                        overrideStyle.textContent = '';
                    }
                }

                // Sync body class for any CSS rules that target body.vscode-light/dark
                body.classList.remove('vscode-light', 'vscode-dark', 'vscode-theme-auto');
                if (currentTheme === 'light') {
                    body.classList.add('vscode-light');
                } else if (currentTheme === 'dark') {
                    body.classList.add('vscode-dark');
                } else {
                    const detectedTheme = detectVSCodeTheme();
                    body.classList.add(detectedTheme || 'vscode-light');
                }

                // Reconnect observer now that our DOM changes are done
                observer.observe(document.body, { attributes: true, attributeFilter: ['class', 'style'] });
                observer.observe(document.documentElement, { attributes: true, attributeFilter: ['class', 'style'] });

                updateThemeButton();
            }

            // Initialize theme
            // Defer slightly so VS Code's own body-class injection (which happens post-load)
            // is detected and overridden when an explicit theme is configured.
            function initializeTheme() {
                applyConfigTheme();
                // Re-apply after a tick to catch VS Code's post-load body class injection
                setTimeout(applyConfigTheme, 0);
            }
            
            // Set up the MutationObserver BEFORE calling initializeTheme(), so that
            // applyConfigTheme() can safely call observer.disconnect()/observe() on first run.
            // (const is in the temporal dead zone until this line executes)
            const observer = new MutationObserver(() => {
                applyConfigTheme();
            });
            
            // Start observing before we call initializeTheme
            observer.observe(document.body, {
                attributes: true,
                attributeFilter: ['class', 'style']
            });
            observer.observe(document.documentElement, {
                attributes: true,
                attributeFilter: ['class', 'style']
            });

            // Initialize theme on load (observer is now ready)
            initializeTheme();

            // Split the document into pages of the document's own size (.docx-page, set by
            // DocxHandler), as Word does: a paragraph that does not fit is split at the first
            // line that would cross the bottom margin and continues on the next page; other
            // blocks (tables, lists, images, display equations) move whole. A block taller than
            // a page gets a page of its own, which grows. Run again once images have loaded.
            function paginate() {
                const doc = document.getElementById('document');
                const pages = doc ? Array.from(doc.querySelectorAll(':scope > .docx-page')) : [];
                if (pages.length === 0) {
                    return;
                }
                doc.querySelectorAll('.cm-card').forEach(c => c.remove());   // placed again after
                const first = pages[0];
                for (let i = 1; i < pages.length; i++) {          // gather everything back
                    while (pages[i].firstChild) {
                        first.appendChild(pages[i].firstChild);
                    }
                    pages[i].remove();
                }
                first.querySelectorAll(':scope > [data-cont]').forEach(c => {   // re-join split paragraphs
                    const head = c.previousElementSibling;
                    if (head) {
                        while (c.firstChild) {
                            head.appendChild(c.firstChild);
                        }
                        head.style.marginBottom = '';
                        head.normalize();
                        c.remove();
                    }
                });
                const cs = getComputedStyle(first);
                const padTop = parseFloat(cs.paddingTop), padBottom = parseFloat(cs.paddingBottom);
                const avail = parseFloat(cs.minHeight) - padTop - padBottom;
                if (!(avail > 0)) {
                    return;
                }
                // Page breaks: Word's own (saved with its last layout) when the document has
                // them, then no measuring at all; otherwise manual breaks plus measuring.
                const wordBreaks = !!first.querySelector('.pb[data-k="r"]');
                const breakSel = wordBreaks ? '.pb[data-k="r"]' : '.pb[data-k="h"]';
                const queue = Array.from(first.childNodes);
                queue.forEach(n => n.remove());
                let page = first;
                const newPage = () => {
                    const next = document.createElement('div');
                    next.className = 'docx-page';
                    page.after(next);
                    page = next;
                };
                const bottom = el => el.offsetTop + el.offsetHeight - padTop;
                while (queue.length) {
                    const n = queue.shift();
                    page.appendChild(n);
                    if (n.nodeType !== 1) {
                        continue;
                    }
                    const brk = n.tagName === 'TABLE' ? null : breakAtMarker(n, breakSel, page);
                    if (brk === 'before') {                       // break before this block
                        n.remove();
                        newPage();
                        queue.unshift(n);
                        continue;
                    }
                    if (brk) {                                    // break inside: first part stays
                        page.insertBefore(brk, n);
                        n.remove();
                        n.setAttribute('data-cont', '1');
                        n.style.marginTop = '0';
                        brk.style.marginBottom = '0';
                        newPage();
                        queue.unshift(n);
                        continue;
                    }
                    if (wordBreaks || bottom(n) <= avail) {
                        continue;
                    }
                    const head = n.tagName === 'P' ? splitParagraph(n, page, padTop + avail) : null;
                    if (head) {                                   // first lines stay, rest goes on
                        page.insertBefore(head, n);
                        n.remove();
                        n.setAttribute('data-cont', '1');
                        n.style.marginTop = '0';
                        head.style.marginBottom = '0';
                        newPage();
                        queue.unshift(n);
                    } else if (page.childElementCount > 1) {      // move the whole block
                        n.remove();
                        newPage();
                        queue.unshift(n);
                    }
                }
            }

            // First page-break marker in block n with content before it: returns that content
            // as a new element (the rest stays in n), 'before' when the marker is at the start
            // of n and n is not the first block of the page, or null. Markers at the very start
            // of a page are already satisfied; markers inside tables are ignored.
            function breakAtMarker(n, sel, page) {
                for (const m of n.querySelectorAll(sel)) {
                    if (m.closest('table')) {
                        continue;
                    }
                    const r = document.createRange();
                    r.setStart(n, 0);
                    r.setEndBefore(m);
                    const frag = r.cloneContents();
                    if (frag.textContent.trim() !== '' || frag.querySelector('img,math')) {
                        const head = n.cloneNode(false);
                        head.removeAttribute('id');
                        head.appendChild(r.extractContents());
                        return head;
                    }
                    if (page.childElementCount > 1) {
                        return 'before';
                    }
                }
                return null;
            }

            // Split paragraph p (on page) before the first line that ends below y = limit (page
            // coordinates, unscaled). Returns the first part as a new element and leaves the rest
            // in p, or null when not even the first line fits. Equations are never split.
            function splitParagraph(p, page, limit) {
                const pr = page.getBoundingClientRect();
                const scale = pr.height / page.offsetHeight || 1;   // zoom (CSS transform)
                const limitY = pr.top + limit * scale;
                const walker = document.createTreeWalker(p, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
                    acceptNode: n => n.nodeType === 1
                        ? (n.localName === 'math' ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_SKIP)
                        : NodeFilter.FILTER_ACCEPT
                });
                const range = document.createRange();
                const charBottom = (t, i) => {
                    range.setStart(t, i);
                    range.setEnd(t, i + 1);
                    const r = range.getClientRects();
                    return r.length ? r[r.length - 1].bottom : -Infinity;   // collapsed space: no rect
                };
                let node, at = -1;
                while ((node = walker.nextNode())) {
                    const len = node.data.length;
                    if (len === 0 || charBottom(node, len - 1) <= limitY && !/\\s$/.test(node.data)) {
                        continue;                                 // whole text node fits
                    }
                    let lo = 0, hi = len;                         // first char below the limit
                    while (lo < hi) {
                        const mid = (lo + hi) >> 1;
                        if (charBottom(node, mid) > limitY) { hi = mid; } else { lo = mid + 1; }
                    }
                    if (lo < len) {
                        at = lo;
                        break;
                    }
                }
                if (at < 0 || !node) {
                    return null;
                }
                while (at > 0 && !/\\s/.test(node.data[at - 1])) {   // back to the start of the word
                    at--;
                }
                const cut = document.createRange();
                cut.setStart(p, 0);
                cut.setEnd(node, at);
                if (cut.toString().trim() === '') {
                    return null;                                  // first line does not fit
                }
                const head = p.cloneNode(false);
                head.removeAttribute('id');
                head.appendChild(cut.extractContents());
                return head;
            }
            // Word comments (JSON from DocxHandler, anchors <span class="cm-s|cm-e|cm-r">):
            // the commented text is highlighted (CSS Custom Highlight API, no DOM change) and
            // each comment, with its replies, is a card beside its line in the column right of
            // the page; cards on a page are stacked so they do not overlap. Click a card to
            // emphasise its text.
            let commentRanges = {};
            function renderComments() {
                const doc = document.getElementById('document');
                const dataEl = document.getElementById('docx-comments');
                if (!doc || !dataEl) {
                    return;
                }
                doc.querySelectorAll('.cm-card').forEach(c => c.remove());
                if (doc.classList.contains('no-comments')) {      // clean view
                    doc.classList.remove('has-comments');
                    if (typeof CSS !== 'undefined' && CSS.highlights) {
                        CSS.highlights.delete('docx-comment');
                        CSS.highlights.delete('docx-comment-active');
                    }
                    return;
                }
                let list = [];
                try {
                    list = JSON.parse(dataEl.textContent || '[]');
                } catch (e) {
                    return;
                }
                if (!list.length) {
                    return;
                }
                doc.classList.add('has-comments');
                const byId = {};
                list.forEach(c => { byId[c.id] = c; c.replies = []; });
                list.forEach(c => { if (c.parent && byId[c.parent]) { byId[c.parent].replies.push(c); } });
                const fmt = d => {
                    const t = d ? new Date(d) : null;
                    return t && !isNaN(t) ? t.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) : '';
                };
                const entry = (c, cls) => {
                    const box = document.createElement('div');
                    box.className = cls;
                    const head = document.createElement('div');
                    head.className = 'cm-head';
                    head.textContent = c.author || c.initials || 'Comment';
                    const date = document.createElement('span');
                    date.className = 'cm-date';
                    date.textContent = fmt(c.date);
                    head.appendChild(date);
                    if (c.done && cls === 'cm-main') {
                        const tag = document.createElement('span');
                        tag.className = 'cm-tag';
                        tag.textContent = 'Resolved';
                        head.appendChild(tag);
                    }
                    box.appendChild(head);
                    (c.paras || []).forEach(t => {
                        const p = document.createElement('p');
                        p.textContent = t;
                        box.appendChild(p);
                    });
                    return box;
                };
                const sel = (cls, id) => doc.querySelector('.' + cls + '[data-cid="' + id + '"]');
                const placed = [];
                const all = [];
                commentRanges = {};
                list.forEach(c => {
                    if (c.parent && byId[c.parent]) {
                        return;                                   // shown inside its parent's card
                    }
                    const s = sel('cm-s', c.id), e = sel('cm-e', c.id), r = sel('cm-r', c.id);
                    const anchor = s || r || e;
                    const page = anchor ? anchor.closest('.docx-page') : null;
                    if (!page) {
                        return;
                    }
                    if (s && e) {
                        try {
                            const range = document.createRange();
                            range.setStartAfter(s);
                            range.setEndBefore(e);
                            if (!range.collapsed) {
                                commentRanges[c.id] = range;
                                if (!c.done) {
                                    all.push(range);              // resolved: highlighted only when clicked
                                }
                            }
                        } catch (err) { /* malformed range: card only */ }
                    }
                    const card = document.createElement('div');
                    card.className = 'cm-card' + (c.done ? ' done' : '');
                    card.dataset.cid = c.id;
                    card.appendChild(entry(c, 'cm-main'));
                    c.replies.forEach(rep => card.appendChild(entry(rep, 'cm-reply')));
                    card.addEventListener('click', () => activateComment(c.id));
                    const pr = page.getBoundingClientRect();
                    const scale = pr.height / page.offsetHeight || 1;
                    // y in document coordinates (all pages share #document as offset parent)
                    placed.push({ card, page, y: page.offsetTop + (anchor.getBoundingClientRect().top - pr.top) / scale });
                    page.appendChild(card);
                });
                if (typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined') {
                    CSS.highlights.set('docx-comment', new Highlight(...all));
                }
                commentCards = placed.sort((a, b) => a.y - b.y);
                stackComments();
            }
            // Cards top to bottom through the whole document, each at its line or just below
            // the card before it (cards may run past the end of their page).
            let commentCards = [];
            function stackComments() {
                let lastBottom = -Infinity;
                commentCards.forEach(({ card, page, y }) => {
                    const top = Math.max(y, lastBottom + 6);
                    card.style.top = (top - page.offsetTop) + 'px';
                    lastBottom = top + card.offsetHeight;
                });
            }
            function activateComment(id) {
                document.querySelectorAll('.cm-card.active').forEach(c => c.classList.remove('active'));
                const card = document.querySelector('.cm-card[data-cid="' + id + '"]');
                if (card) {
                    card.classList.add('active');                 // resolved cards open up
                }
                stackComments();
                if (typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined') {
                    const r = commentRanges[id];
                    if (r) {
                        CSS.highlights.set('docx-comment-active', new Highlight(r));
                    } else {
                        CSS.highlights.delete('docx-comment-active');
                    }
                }
            }
            function layout() {
                paginate();
                renderComments();
            }
            // View, cycled by the button: markup (tracked changes as insertions / deletions,
            // comments), final (final text, comments), clean (final text only, no comments).
            const markupBtn = document.getElementById('toggleMarkup');
            if (markupBtn) {
                const views = [
                    { cls: [], icon: '±', title: 'Markup: tracked changes and comments (click: final)' },
                    { cls: ['tc-final'], icon: '✓', title: 'Final: final text with comments (click: clean)' },
                    { cls: ['tc-final', 'no-comments'], icon: '◻', title: 'Clean: final text only (click: markup)' },
                ];
                let view = 0;
                if (!document.querySelector('#document .tc') && !document.getElementById('docx-comments')) {
                    markupBtn.style.display = 'none';             // nothing to show or hide
                }
                markupBtn.title = views[0].title;
                markupBtn.addEventListener('click', () => {
                    const doc = document.getElementById('document');
                    view = (view + 1) % views.length;
                    doc.classList.remove('tc-final', 'no-comments');
                    views[view].cls.forEach(c => doc.classList.add(c));
                    markupBtn.textContent = views[view].icon;
                    markupBtn.title = views[view].title;
                    layout();                                     // heights changed
                });
            }
            layout();
            window.addEventListener('load', layout);
            
            function updateThemeButton() {
                const button = document.getElementById('themeToggle');
                if (currentTheme === 'auto') {
                    button.textContent = '🔄';
                    button.title = 'Theme: Auto (Click for Dark)';
                } else if (currentTheme === 'light') {
                    button.textContent = '🌙';
                    button.title = 'Theme: Light (Click for Auto)';
                } else {
                    button.textContent = '☀️';
                    button.title = 'Theme: Dark (Click for Light)';
                }
            }
            
            // Zoom controls
            document.getElementById('zoomIn').addEventListener('click', () => {
                if (currentZoom < 3.0) {
                    currentZoom = Math.min(3.0, currentZoom + 0.25);
                    updateZoom();
                    // Send message back to extension
                    vscode.postMessage({
                        command: 'zoomChanged',
                        zoom: currentZoom
                    });
                }
            });
            
            document.getElementById('zoomOut').addEventListener('click', () => {
                if (currentZoom > 0.5) {
                    currentZoom = Math.max(0.5, currentZoom - 0.25);
                    updateZoom();
                    // Send message back to extension
                    vscode.postMessage({
                        command: 'zoomChanged',
                        zoom: currentZoom
                    });
                }
            });
            
            document.getElementById('resetZoom').addEventListener('click', () => {
                currentZoom = 1.0;
                updateZoom();
                // Send message back to extension
                vscode.postMessage({
                    command: 'zoomChanged',
                    zoom: currentZoom
                });
            });
            
            function updateZoom() {
                const doc = document.getElementById('document');
                const zoomLevel = document.getElementById('zoomLevel');
                
                // Apply zoom using CSS transform directly
                doc.style.transform = 'scale(' + currentZoom + ')';
                
                const zoomPercent = Math.round(currentZoom * 100);
                zoomLevel.textContent = zoomPercent + '%';
            }
            
            // Outline toggle
            document.getElementById('toggleOutline').addEventListener('click', () => {
                outlineVisible = !outlineVisible;
                const outline = document.getElementById('outlinePanel');
                const button = document.getElementById('toggleOutline');
                
                outline.classList.toggle('hidden', !outlineVisible);
                button.textContent = outlineVisible ? '◧' : '◨';
                
                // Send message back to extension
                vscode.postMessage({
                    command: 'outlineToggled',
                    visible: outlineVisible
                });
            });
            
            // Theme toggle
            document.getElementById('themeToggle').addEventListener('click', () => {
                if (currentTheme === 'dark') {
                    currentTheme = 'light';
                } else if (currentTheme === 'light') {
                    currentTheme = 'auto';
                } else {
                    // If auto, switch to dark
                    currentTheme = 'dark';
                }
                
                updateTheme();
                
                // Send message back to extension
                vscode.postMessage({
                    command: 'themeChanged',
                    theme: currentTheme
                });
            });
            
            function updateTheme() {
                // Delegate to applyConfigTheme() which enforces the theme correctly
                // and handles the MutationObserver guard
                applyConfigTheme();
            }
            
            // Toolbar toggle
            document.getElementById('hideToolbar').addEventListener('click', () => {
                toolbarVisible = false;
                updateToolbarVisibility();
                
                // Send message back to extension
                vscode.postMessage({
                    command: 'toolbarToggled',
                    visible: toolbarVisible
                });
            });
            
            document.getElementById('showToolbar').addEventListener('click', () => {
                toolbarVisible = true;
                updateToolbarVisibility();
                
                // Send message back to extension
                vscode.postMessage({
                    command: 'toolbarToggled',
                    visible: toolbarVisible
                });
            });
            
            function updateToolbarVisibility() {
                const mainToolbar = document.getElementById('mainToolbar');
                const miniToolbar = document.getElementById('miniToolbar');
                
                if (toolbarVisible) {
                    mainToolbar.classList.remove('hidden');
                    miniToolbar.classList.add('hidden');
                } else {
                    mainToolbar.classList.add('hidden');
                    miniToolbar.classList.remove('hidden');
                }
            }
            
            // Outline navigation
            document.querySelectorAll('.docx-outline-item').forEach(item => {
                item.addEventListener('click', () => {
                    const targetId = item.getAttribute('data-target');
                    const target = document.getElementById(targetId);
                    
                    if (target) {
                        // Remove active class from all outline items
                        document.querySelectorAll('.docx-outline-item').forEach(i => i.classList.remove('active'));
                        
                        // Add active class to clicked item
                        item.classList.add('active');
                        
                        // Scroll to target
                        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                });
            });
            
            // Search functionality
            document.getElementById('searchBtn').addEventListener('click', () => {
                const searchPanel = document.getElementById('searchPanel');
                searchPanel.classList.toggle('visible');
                if (searchPanel.classList.contains('visible')) {
                    document.getElementById('searchInput').focus();
                }
            });
            
            document.getElementById('closeSearch').addEventListener('click', () => {
                document.getElementById('searchPanel').classList.remove('visible');
                clearSearchResults();
            });
            
            document.getElementById('searchInput').addEventListener('input', performSearch);
            document.getElementById('searchNext').addEventListener('click', () => navigateSearch(1));
            document.getElementById('searchPrev').addEventListener('click', () => navigateSearch(-1));

            function updateSearchCounter() {
                const counter = document.getElementById('searchCount');
                const total = searchResults.length;
                const current = total > 0 && currentSearchIndex >= 0 ? currentSearchIndex + 1 : 0;

                counter.textContent = current + '/' + total;
            }

            function updateSearchControls() {
                const hasResults = searchResults.length > 0;
                document.getElementById('searchNext').disabled = !hasResults;
                document.getElementById('searchPrev').disabled = !hasResults;
                updateSearchCounter();
            }

            function escapeRegExp(value) {
                return value.replace(/[.*+?^\${}()|[\]\\]/g, '\\$&');
            }
            
            function performSearch() {
                const searchTerm = document.getElementById('searchInput').value.trim();
                clearSearchResults();
                
                if (searchTerm.length === 0) {
                    updateSearchControls();
                    return;
                }
                
                const content = document.getElementById('document');
                const walker = document.createTreeWalker(
                    content,
                    NodeFilter.SHOW_TEXT,
                    null,
                    false
                );
                
                const textNodes = [];
                let node;
                while (node = walker.nextNode()) {
                    textNodes.push(node);
                }
                
                searchResults = [];
                const regex = new RegExp(escapeRegExp(searchTerm), 'gi');
                textNodes.forEach(textNode => {
                    const text = textNode.textContent || '';
                    let match;

                    regex.lastIndex = 0;
                    if (!regex.test(text)) {
                        return;
                    }

                    regex.lastIndex = 0;
                    const fragment = document.createDocumentFragment();
                    let lastIndex = 0;
                    
                    while ((match = regex.exec(text)) !== null) {
                        if (match.index > lastIndex) {
                            fragment.appendChild(document.createTextNode(text.substring(lastIndex, match.index)));
                        }

                        const highlightSpan = document.createElement('span');
                        highlightSpan.className = 'search-highlight';
                        highlightSpan.textContent = match[0];
                        fragment.appendChild(highlightSpan);

                        searchResults.push({ element: highlightSpan });
                        lastIndex = match.index + match[0].length;
                    }

                    if (lastIndex < text.length) {
                        fragment.appendChild(document.createTextNode(text.substring(lastIndex)));
                    }

                    textNode.parentNode.replaceChild(fragment, textNode);
                });
                
                if (searchResults.length > 0) {
                    currentSearchIndex = 0;
                    navigateToSearchResult(0);
                } else {
                    currentSearchIndex = -1;
                    updateSearchControls();
                }
            }
            
            function clearSearchResults() {
                document.querySelectorAll('.search-highlight').forEach(highlight => {
                    const parent = highlight.parentNode;
                    parent.replaceChild(document.createTextNode(highlight.textContent), highlight);
                    parent.normalize();
                });
                searchResults = [];
                currentSearchIndex = -1;
                activeSearchElement = null;
                updateSearchControls();
            }
            
            function navigateSearch(direction) {
                if (searchResults.length === 0) return;
                
                currentSearchIndex += direction;
                if (currentSearchIndex >= searchResults.length) currentSearchIndex = 0;
                if (currentSearchIndex < 0) currentSearchIndex = searchResults.length - 1;
                
                navigateToSearchResult(currentSearchIndex);
            }
            
            function navigateToSearchResult(index) {
                if (index < 0 || index >= searchResults.length) return;
                
                const result = searchResults[index];
                if (result.element) {
                    if (activeSearchElement) {
                        activeSearchElement.classList.remove('current');
                    }

                    result.element.classList.add('current');
                    activeSearchElement = result.element;
                    result.element.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    updateSearchControls();
                }
            }

            updateSearchControls();
            
            // Sync Scroll
            const docContent = document.getElementById('documentContent');
            let isSyncing = false;
            
            if (docContent) {
                docContent.addEventListener('scroll', () => {
                    if (!isSyncing) {
                        const maxScroll = docContent.scrollHeight - docContent.clientHeight;
                        if (maxScroll > 0) {
                            const pct = docContent.scrollTop / maxScroll;
                            vscode.postMessage({ command: 'scroll', scrollPercent: pct });
                        }
                    }
                    isSyncing = false;
                });
            }

            // Handle messages from extension
            window.addEventListener('message', event => {
                const message = event.data;
                switch (message.command) {
                    case 'updateZoom':
                        currentZoom = message.zoom;
                        updateZoom();
                        break;
                    case 'toggleOutline':
                        outlineVisible = message.visible;
                        const outline = document.getElementById('outlinePanel');
                        const button = document.getElementById('toggleOutline');
                        outline.classList.toggle('hidden', !outlineVisible);
                        button.textContent = outlineVisible ? '◧' : '◨';
                        break;
                    case 'updateTheme':
                        currentTheme = message.theme;
                        updateTheme();
                        break;
                    case 'toggleToolbar':
                        toolbarVisible = message.visible;
                        updateToolbarVisibility();
                        break;
                    case 'syncScroll':
                        if (docContent) {
                            isSyncing = true;
                            const maxScroll = docContent.scrollHeight - docContent.clientHeight;
                            docContent.scrollTop = message.scrollPercent * maxScroll;
                        }
                        break;
                    case 'highlight':
                        highlightDiffs(message.diffs);
                        break;
                }
            });

            function highlightDiffs(diffs) {
                const doc = document.getElementById('document');
                if (!doc) return;
                
                doc.querySelectorAll('.diff-added, .diff-removed').forEach(el => {
                    el.classList.remove('diff-added', 'diff-removed');
                });

                const elements = Array.from(doc.children);
                
                if (diffs.added) {
                    diffs.added.forEach(idx => {
                        if (elements[idx]) elements[idx].classList.add('diff-added');
                    });
                }
                if (diffs.removed) {
                    diffs.removed.forEach(idx => {
                        if (elements[idx]) elements[idx].classList.add('diff-removed');
                    });
                }
            }
            
            // Keyboard shortcuts
            document.addEventListener('keydown', (e) => {
                if (e.ctrlKey || e.metaKey) {
                    switch (e.key) {
                        case '+':
                        case '=':
                            e.preventDefault();
                            document.getElementById('zoomIn').click();
                            break;
                        case '-':
                            e.preventDefault();
                            document.getElementById('zoomOut').click();
                            break;
                        case '0':
                            e.preventDefault();
                            document.getElementById('resetZoom').click();
                            break;
                        case 'f':
                            e.preventDefault();
                            document.getElementById('searchBtn').click();
                            break;
                        case 'h':
                            e.preventDefault();
                            if (toolbarVisible) {
                                document.getElementById('hideToolbar').click();
                            } else {
                                document.getElementById('showToolbar').click();
                            }
                            break;
                    }
                }
                
                if (e.key === 'Escape') {
                    document.getElementById('closeSearch').click();
                }
            });
        })();
        `;
    }

    public static updateZoom(zoom: number) {
        this.currentZoom = zoom;
    }

    public static toggleOutline() {
        this.outlineVisible = !this.outlineVisible;
    }

    public static updateTheme(theme: string) {
        this.currentTheme = theme;
    }

    public static toggleToolbar() {
        this.toolbarVisible = !this.toolbarVisible;
    }
}

interface OutlineItem {
    level: number;
    text: string;
    id: string;
}

// Legacy function for backward compatibility
export async function renderDocuments(docxPath: string, panel: vscode.WebviewPanel) {
    return DocumentRenderer.renderDocument(vscode.Uri.file(docxPath), panel);
}