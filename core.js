// core.js - Shared theme customizer for ICMS
(function () {
    const DEFAULT_SETTINGS = {
        bgColor: '#e6f4f1',
        bgColor2: '#cceeee',
        sidebarColor: '#0b3c55',
        sidebarColor2: '#17688a',
        heroBgColor: '#0b3c55',
        heroBgColor2: '#0F5A7D',
        tableHeaderBgColor: '#0F5A7D',
        tableHeaderTextColor: '#ffffff',
        accentColor: '#007474',
        sidebarActiveColor: 'rgba(255, 255, 255, 0.16)',
        fontFamily: 'system-ui',
        sidebarFontSize: '12',
        titleFontSize: '23',
        tableFontSize: '11.5',
        tableMinColWidth: '80',
        tableMaxColWidth: '300',
        tableMinRowHeight: '32',
        tableMaxRowHeight: '80',
        cardBorderRadius: '16',
        buttonBorderRadius: '8',
        sidebarWidth: '240'
    };

    function applyCustomStyles() {
        const settingsRaw = localStorage.getItem('icms_settings');
        let settings = DEFAULT_SETTINGS;
        
        if (settingsRaw) {
            try {
                settings = { ...DEFAULT_SETTINGS, ...JSON.parse(settingsRaw) };
            } catch (e) {
                settings = DEFAULT_SETTINGS;
            }
        }

        try {
            let css = '';

            // 1. Background Color / Gradient (tidak diterapkan di login.html agar login-bg.jpg tetap tampil)
            const isLoginPage = window.location.pathname.endsWith('login.html');
            if (settings.bgColor && !isLoginPage) {
                let bgVal = settings.bgColor;
                if (settings.bgColor2) {
                    bgVal = `linear-gradient(135deg, ${settings.bgColor}, ${settings.bgColor2})`;
                }
                css += `
                    body {
                        background: ${bgVal} !important;
                    }
                `;
            }

            // 2. Sidebar Color / Gradient
            if (settings.sidebarColor) {
                let sideVal = settings.sidebarColor;
                if (settings.sidebarColor2) {
                    sideVal = `linear-gradient(180deg, ${settings.sidebarColor}, ${settings.sidebarColor2})`;
                }
                css += `
                    .sidebar {
                        background: ${sideVal} !important;
                    }
                `;
            }

            // 2b. Sidebar Width (tanpa !important agar fleksibel)
            if (settings.sidebarWidth) {
                const w = parseInt(settings.sidebarWidth);
                css += `
                    .sidebar {
                        width: ${w}px !important;
                    }
                    .main {
                        margin-left: ${w}px !important;
                    }
                `;
            }

            // 2b. Top Panel (Hero / Filter Card) Color / Gradient
            if (settings.heroBgColor) {
                let heroVal = settings.heroBgColor;
                if (settings.heroBgColor2) {
                    heroVal = `linear-gradient(90deg, ${settings.heroBgColor}, ${settings.heroBgColor2})`;
                }
                css += `
                    .hero {
                        background: ${heroVal} !important;
                    }
                `;
            }

            // 3. General Font Family
            if (settings.fontFamily) {
                if (settings.fontFamily !== 'system-ui') {
                    let fontLink = document.getElementById('custom-font-link');
                    const fontHref = `https://fonts.googleapis.com/css2?family=${settings.fontFamily.replace(/\s+/g, '+')}:wght@300;400;500;600;700&display=swap`;
                    if (!fontLink) {
                        fontLink = document.createElement('link');
                        fontLink.id = 'custom-font-link';
                        fontLink.rel = 'stylesheet';
                        document.head.appendChild(fontLink);
                    }
                    if (fontLink.href !== fontHref) {
                        fontLink.href = fontHref;
                    }
                } else {
                    const fontLink = document.getElementById('custom-font-link');
                    if (fontLink) fontLink.remove();
                }
                css += `
                    body, input, select, button, textarea, .brand-title, .brand-subtitle, th, td, a, p, span, div {
                        font-family: '${settings.fontFamily}', system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif !important;
                    }
                `;
            }

            // 4. Sidebar Font Size (Tanpa !important agar dapat dioverride oleh inline style / menu edit)
            if (settings.sidebarFontSize) {
                css += `
                    .sidebar, .sidebar nav a, .sidebar-bottom, .sidebar-bottom button, .sidebar-bottom span {
                        font-size: ${settings.sidebarFontSize}px;
                    }
                `;
            }

            // 4b. Title / Header Font Size (Tanpa !important)
            if (settings.titleFontSize) {
                css += `
                    h1, .hero h1, .preview-title, .section-title, .brand-title {
                        font-size: ${settings.titleFontSize}px;
                    }
                `;
            }

            // 4c. Table Font Size (Tanpa !important)
            if (settings.tableFontSize) {
                css += `
                    table th, table td, table tr, th, td, tr, .preview-table th, .preview-table td {
                        font-size: ${settings.tableFontSize}px;
                    }
                `;
            }

            // 4d. Table Header Colors (Tanpa !important agar bisa ditimpa saat edit)
            if (settings.tableHeaderBgColor || settings.tableHeaderTextColor) {
                css += `
                    table th, th, .preview-table th {
                `;
                if (settings.tableHeaderBgColor) {
                    css += `
                        background: ${settings.tableHeaderBgColor};
                        background-color: ${settings.tableHeaderBgColor};
                    `;
                }
                if (settings.tableHeaderTextColor) {
                    css += `
                        color: ${settings.tableHeaderTextColor};
                    `;
                }
                css += `
                    }
                `;
            }

            // 5. Table Column & Row Sizes
            if (settings.tableMinColWidth || settings.tableMaxColWidth) {
                css += `
                    table th, table td {
                `;
                if (settings.tableMinColWidth) {
                    css += `min-width: ${settings.tableMinColWidth}px !important;`;
                }
                if (settings.tableMaxColWidth) {
                    css += `max-width: ${settings.tableMaxColWidth}px !important;`;
                }
                css += `
                    }
                `;
            }

            if (settings.tableMinRowHeight || settings.tableMaxRowHeight) {
                css += `
                    table tr {
                `;
                if (settings.tableMinRowHeight) {
                    css += `height: ${settings.tableMinRowHeight}px !important;`;
                }
                css += `
                    }
                    table th, table td {
                `;
                if (settings.tableMinRowHeight) {
                    css += `height: ${settings.tableMinRowHeight}px !important;`;
                    // Adjust padding to center text vertically if rows are tall
                    const padValue = Math.max(2, Math.floor((parseInt(settings.tableMinRowHeight) - 20) / 2));
                    css += `padding-top: ${padValue}px !important; padding-bottom: ${padValue}px !important;`;
                }
                if (settings.tableMaxRowHeight) {
                    css += `max-height: ${settings.tableMaxRowHeight}px !important; overflow: hidden !important;`;
                }
                css += `
                    }
                `;
            }

            // 6. Accent Color
            if (settings.accentColor) {
                css += `
                    :root {
                        --accent: ${settings.accentColor} !important;
                        --border: ${settings.accentColor} !important;
                    }
                    .brand-mark, .primary, .btn.cyan, button[type="submit"] {
                        background: ${settings.accentColor} !important;
                        color: #ffffff !important;
                    }
                `;
            }

            // 7. Sidebar Active Color
            if (settings.sidebarActiveColor) {
                css += `
                    .sidebar nav a.active, .sidebar nav a:hover {
                        background: ${settings.sidebarActiveColor} !important;
                        color: #ffffff !important;
                    }
                `;
            }

            // 8. Kelengkungan Sudut (Border Radius)
            if (settings.cardBorderRadius !== undefined) {
                css += `
                    .card, .hero, .modal-box, .preview-shell, .preview-card {
                        border-radius: ${settings.cardBorderRadius}px !important;
                    }
                `;
            }
            if (settings.buttonBorderRadius !== undefined) {
                css += `
                    .btn, button, input, select, textarea, .brand-mark, .preview-brand, .preview-nav-item {
                        border-radius: ${settings.buttonBorderRadius}px !important;
                    }
                `;
            }

            // Inject styles
            let styleEl = document.getElementById('custom-theme-styles');
            if (!styleEl) {
                styleEl = document.createElement('style');
                styleEl.id = 'custom-theme-styles';
            }
            styleEl.innerHTML = css;
            document.head.appendChild(styleEl);

        } catch (e) {
            console.error('Error applying custom styles:', e);
        }
    }

    // Run immediately to avoid flashing
    applyCustomStyles();

    // Run again on DOMContentLoaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', applyCustomStyles);
    }

    // Expose globally
    window.applyCustomStyles = applyCustomStyles;
})();
