# SheetGPT

Use ChatGPT directly in your Google Sheets! Ask questions, analyze data, or get AI assistance right from your spreadsheet cells.

## Quick Start Guide

### Step 1: Create Your OpenAI Account (Skip if you already have one)
1. Visit [OpenAI's website](https://platform.openai.com/signup)
2. Create an account or sign in
3. Go to [API Keys page](https://platform.openai.com/api-keys)
4. Click "Create new secret key"
5. Copy your API key (you'll need this later!)

### Step 2: Set Up in Google Sheets
1. Open a new Google Sheet at [sheets.new](https://sheets.new)
2. Click on `Extensions` in the top menu
3. Select `Apps Script` - this will open a new tab

### Step 3: Add the Code
1. In the Apps Script tab, you'll see a file named `Code.gs`
2. Delete any existing code in this file
3. Copy all the code from the `code.js` file in this repository
4. Paste it into your `Code.gs`
5. Click the 💾 Save icon (or press Ctrl/Cmd + S)

### Step 4: Authorize the Script
1. In the Apps Script editor, select `onOpen` from the dropdown menu next to the "Debug" button
2. Click the ▶️ "Run" button
3. Click "Review Permissions" in the popup
4. Choose your Google account
5. You might see a warning because the app isn't verified - click "Advanced" and then "Go to SheetGPT (unsafe)"
6. Click "Allow"

### Step 5: Add Your OpenAI Key
1. Go back to your Google Sheet
2. You'll see a new menu item called `SheetGPT`
3. Click `SheetGPT` > `Set Secret Key`
4. Paste your OpenAI API key from Step 1
5. Click "OK"

### Step 6: Start Using SheetGPT! 🎉
Type in any cell:
```
=SheetGPT("your question here")
```

Examples:
- `=SheetGPT("Summarize this data", A2)`
- `=SheetGPT("Translate this to Spanish", A2)`
- `=SheetGPT("What insights can you give me about these numbers?", A2:B40)`

## Tips
- The formula can reference other cells: `=SheetGPT(A1)`
- You can select multiple cells: `=SheetGPT(A1:B10)`
- Each query uses your OpenAI API credits
- Responses are cached in the cell until you refresh

## Troubleshooting
- If you see `#NO_API_KEY`, go to `SheetGPT > Set Secret Key` to add your API key
- If you see `#NO_INPUT`, make sure you've provided some text or cell reference
- If the menu doesn't appear, refresh the page


## Why use this?

Not everything needs to be an app. 

What I love most about google sheets is that it's easy for anyone to use.

This formula is free and open source. I could have made it a paid product. but it is too generalised to charge anything for it. In order to make it a paid product I think I'll have to make it specialised into a niche problem solving tool. 

also, I believe AI will be more and more commoditized in coming days. 

so, since this is a generalised, tool you should be able to enter in your own API keys and use it on your dime.

feel free to fork this repo and make your own version of it. 

Cheers!
Ali Khundmiri