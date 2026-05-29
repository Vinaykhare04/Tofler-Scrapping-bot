# Playwright Browser Installation Guide

## Issue
```
BrowserType.launch: Executable doesn't exist at ...
Looks like Playwright was just installed or updated.
Please run the following command to download new browsers:
playwright install
```

## Quick Fix

### Windows
```bash
python -m playwright install chromium
# Or run the batch file:
setup_playwright.bat
```

### macOS/Linux
```bash
python -m playwright install chromium
# Or run the shell script:
bash setup_playwright.sh
```

## For Streamlit Cloud Deployment

The app now has **automatic browser installation** built-in. However, you can also manually set it up:

### Method 1: Auto-Installation (Recommended)
The app will automatically install browsers on first run:
- No additional setup needed
- Happens in background during first app load
- Creates `.cache/ms-playwright` folder automatically

### Method 2: Pre-Install (Faster startup)
Create a `.streamlit/post_app_commands.txt` file in your GitHub repo:

```txt
python -m playwright install chromium
```

Then in your GitHub repo root, create `.streamlit/config.toml`:
```toml
[client]
toolbarMode = "minimal"
```

## Files Added

- **packages.txt** - System dependencies for Streamlit Cloud
- **setup_playwright.sh** - Linux/macOS browser installation script
- **setup_playwright.bat** - Windows browser installation script
- **app.py (updated)** - Now includes auto-install function

## Verification

After installation, verify browsers are installed:

```bash
# Should show installed browser paths
python -c "from playwright.sync_api import sync_playwright; sync_playwright().start()"
```

## What's Inside packages.txt?

This file tells Streamlit Cloud to install system libraries required by Playwright:
- Chromium dependencies (X11, fonts, libraries)
- Used automatically by Streamlit Cloud
- No manual action needed

## Troubleshooting

### Still getting "Executable doesn't exist"?
1. **Local**: Run `python -m playwright install chromium` in your project directory
2. **Streamlit Cloud**: 
   - Refresh the app page (F5)
   - Wait 2-3 minutes for background installation
   - Check Streamlit Cloud logs for installation progress

### Slow first load?
This is normal on first deployment. The app:
1. Installs Playwright package (pip)
2. Installs system dependencies (apt-get)
3. Downloads Chromium browser (~300MB)

Total time: ~2-5 minutes on first deployment. Subsequent runs are instant.

### Storage full error?
Chromium takes ~500MB. Check Streamlit Cloud resource limits:
- Storage: 1GB per app
- Memory: 1GB

If you exceed limits, consider:
- Removing old deployment logs
- Using a different cloud service with more resources
- Running scraper on separate backend (not in Streamlit Cloud)

## Architecture

```
Local Development:
app.py → ensure_browsers_installed() → playwright install → run scraper

Streamlit Cloud:
app.py → ensure_browsers_installed() → playwright install (cached) → run scraper
packages.txt → system dependencies (installed by Streamlit)
```

## Next Steps

✅ **For local testing:**
```bash
python -m playwright install chromium
streamlit run app.py
```

✅ **For Streamlit Cloud:**
1. Push to GitHub
2. Deploy to share.streamlit.io
3. Wait for browsers to install (first load)
4. App will work automatically

---

**All fixes applied automatically. No manual intervention needed!** 🚀
