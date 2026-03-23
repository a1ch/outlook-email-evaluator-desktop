# Outlook Email Evaluator - Desktop Add-in

AI-powered spam and phishing detection for desktop Outlook (Windows/Mac) and Outlook on the web, using Claude AI.

> Ported from the [Chrome extension](https://github.com/a1ch/outlook-email-evaluator) to an Office Add-in.

## Files
- `manifest.xml` - Office Add-in manifest (points to GitHub Pages)
- `taskpane.html` - Main UI
- `taskpane.js` - All logic: email reading, link extraction, Claude API call
- `taskpane.css` - Styles
- `commands.html` - Required Office boilerplate

## Quick Start
1. Push this repo to GitHub
2. Enable GitHub Pages: Settings -> Pages -> Deploy from main branch / root
3. In Outlook: Home tab -> Get Add-ins -> My Add-ins -> Upload My Add-in -> select manifest.xml
4. Open any email, click "Analyze Email" in the ribbon
5. Click gear icon, enter your Anthropic API key and org domain

## Key difference from Chrome extension
Instead of scraping the DOM, this uses `Office.context.mailbox.item` to read email content cleanly.
Works in desktop Outlook (Windows & Mac) and Outlook Web.

## License
MIT
