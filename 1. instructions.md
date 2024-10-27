# Project Overview
GPT extension for Google Sheets using Google Apps Script

# Core functionalities
- Connect Google Sheets Apps Script to OpenAI API
- Add a Custom Menu "GPT Extension" in Google Sheets
- Set Up the Sidebar UI that can provide a user-friendly interface where users enter prompts and view results
- There should be a settings panel, so one can edit 
the following parameters:
-- base url of openai compatible model
-- model itself (either chose from the short list of 4 most popular openai models or enter manually the name of the model)
-- temperature (from 0 to 1, of possible use progress bar or something like this)
-- max tokens (from 150 till infinity)

# Documentation

## Connect Google Sheets Apps Script to OpenAI API

```
// Example OpenAI API connection code

// Store your API key securely in Script Properties
function setApiKey() {
  const apiKey = 'your-api-key-here';
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
}

// Get API key from Script Properties
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
}

// Function to call OpenAI API
function callOpenAI(prompt) {
  const apiKey = getApiKey();
  const apiUrl = 'https://api.openai.com/v1/chat/completions';
  
  const requestBody = {
    model: 'gpt-3.5-turbo',
    messages: [
      { role: 'user', content: prompt }
    ],
    temperature: 0.7,
    max_tokens: 1000
  };

  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200) {
      return jsonResponse.choices[0].message.content;
    } else {
      throw new Error(`API Error: ${jsonResponse.error.message}`);
    }
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return `Error: ${error.message}`;
  }
}

// Example usage
function testOpenAI() {
  const result = callOpenAI('What is the capital of France?');
  Logger.log(result);
}
```

## Add a Custom Menu "GPT Extension" in Google Sheets
```
// Example code for creating custom menu in Google Sheets

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GPT Extension')
    .addItem('Show Assistant', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Tools')
      .addItem('Analyze Selected Cells', 'analyzeSelection')
      .addItem('Generate Summary', 'generateSummary'))
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addToUi();
}

// Function to show sidebar (will be implemented later)
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('GPT Assistant')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Example function to analyze selected cells
function analyzeSelection() {
  const selection = SpreadsheetApp.getActiveRange();
  if (!selection) {
    SpreadsheetApp.getUi().alert('Please select some cells first.');
    return;
  }
  
  const values = selection.getValues();
  const prompt = `Analyze this data: ${JSON.stringify(values)}`;
  const analysis = callOpenAI(prompt); // Using the previously defined OpenAI function
  
  // Show results in a modal dialog
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<div style="padding: 20px">${analysis}</div>`),
    'Analysis Results'
  );
}

// Example function to generate summary
function generateSummary() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const prompt = `Generate a summary of this spreadsheet data: ${JSON.stringify(values)}`;
  const summary = callOpenAI(prompt);
  
  // Show results in a modal dialog
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<div style="padding: 20px">${summary}</div>`),
    'Sheet Summary'
  );
}

// Function to show settings dialog
function showSettings() {
  const html = HtmlService.createHtmlOutput(`
    <div style="padding: 20px">
      <h2>GPT Extension Settings</h2>
      <p>API Key: <input type="password" id="apiKey" value="********"></p>
      <button onclick="saveSettings()">Save</button>
      <script>
        function saveSettings() {
          const apiKey = document.getElementById('apiKey').value;
          google.script.run
            .withSuccessHandler(() => alert('Settings saved!'))
            .setApiKey(apiKey);
        }
      </script>
    </div>
  `)
  .setWidth(400)
  .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}
```

## Set Up the Sidebar UI that can provide a user-friendly interface where users enter prompts and view results
Below is an example of a clean and user-friendly sidebar UI implementation. This code should go in your Sidebar.html file.
```
<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <style>
        body {
            font-family: Arial, sans-serif;
            padding: 12px;
            color: #333;
        }

        .container {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .input-section {
            display: flex;
            flex-direction: column;
            gap: 8px;
        }

        textarea {
            width: 100%;
            min-height: 100px;
            padding: 8px;
            border: 1px solid #ccc;
            border-radius: 4px;
            resize: vertical;
            font-family: inherit;
        }

        button {
            background-color: #4285f4;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 500;
        }

        button:hover {
            background-color: #3574e2;
        }

        button:disabled {
            background-color: #ccc;
            cursor: not-allowed;
        }

        .output-section {
            margin-top: 16px;
            border-top: 1px solid #eee;
            padding-top: 16px;
        }

        #output {
            white-space: pre-wrap;
            background-color: #f8f9fa;
            padding: 12px;
            border-radius: 4px;
            min-height: 50px;
        }

        .loading {
            display: none;
            color: #666;
            font-style: italic;
        }

        .error {
            color: #d93025;
            padding: 8px;
            background-color: #fce8e6;
            border-radius: 4px;
            display: none;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="input-section">
            <label for="prompt">Enter your prompt:</label>
            <textarea id="prompt" placeholder="Ask GPT anything about your spreadsheet data..."></textarea>
            <div class="button-group">
                <button onclick="sendPrompt()" id="sendButton">Send</button>
                <button onclick="clearAll()" id="clearButton">Clear</button>
            </div>
        </div>

        <div id="loading" class="loading">Processing your request...</div>
        <div id="error" class="error"></div>

        <div class="output-section">
            <label>Response:</label>
            <div id="output"></div>
        </div>
    </div>

    <script>
        // Get DOM elements
        const promptInput = document.getElementById('prompt');
        const sendButton = document.getElementById('sendButton');
        const clearButton = document.getElementById('clearButton');
        const outputDiv = document.getElementById('output');
        const loadingDiv = document.getElementById('loading');
        const errorDiv = document.getElementById('error');

        // Send prompt to Google Apps Script
        function sendPrompt() {
            const prompt = promptInput.value.trim();
            
            if (!prompt) {
                showError('Please enter a prompt');
                return;
            }

            // Show loading state
            setLoading(true);
            hideError();
            
            // Disable buttons during processing
            sendButton.disabled = true;
            clearButton.disabled = true;

            // Call server-side function
            google.script.run
                .withSuccessHandler(handleSuccess)
                .withFailureHandler(handleError)
                .callOpenAI(prompt);
        }

        // Handle successful response
        function handleSuccess(result) {
            setLoading(false);
            sendButton.disabled = false;
            clearButton.disabled = false;
            outputDiv.textContent = result;
        }

        // Handle error
        function handleError(error) {
            setLoading(false);
            sendButton.disabled = false;
            clearButton.disabled = false;
            showError(error.message || 'An error occurred');
        }

        // Clear all inputs and outputs
        function clearAll() {
            promptInput.value = '';
            outputDiv.textContent = '';
            hideError();
        }

        // Show/hide loading state
        function setLoading(isLoading) {
            loadingDiv.style.display = isLoading ? 'block' : 'none';
        }

        // Show error message
        function showError(message) {
            errorDiv.textContent = message;
            errorDiv.style.display = 'block';
        }

        // Hide error message
        function hideError() {
            errorDiv.style.display = 'none';
        }

        // Handle Enter key to send prompt
        promptInput.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && e.ctrlKey) {
                sendPrompt();
            }
        });
    </script>
</body>
</html>
```
## Model Settings Panel
Here is a code example of model settings panel
```
<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <style>
        body {
            font-family: Arial, sans-serif;
            padding: 20px;
            color: #333;
        }
        .settings-container {
            display: flex;
            flex-direction: column;
            gap: 16px;
            max-width: 500px;
        }
        .setting-group {
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        label {
            font-weight: 500;
        }
        input[type="text"], 
        input[type="number"],
        select {
            padding: 8px;
            border: 1px solid #ccc;
            border-radius: 4px;
            width: 100%;
        }
        .model-input {
            display: none;
        }
        .model-input.show {
            display: block;
        }
        .range-container {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        input[type="range"] {
            flex-grow: 1;
        }
        .value-display {
            min-width: 40px;
        }
        button {
            background-color: #4285f4;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 500;
            margin-top: 16px;
        }
        button:hover {
            background-color: #3574e2;
        }
        .status {
            display: none;
            padding: 8px;
            border-radius: 4px;
            margin-top: 16px;
        }
        .success {
            background-color: #e6f4ea;
            color: #137333;
        }
        .error {
            background-color: #fce8e6;
            color: #d93025;
        }
    </style>
</head>
<body>
    <div class="settings-container">
        <div class="setting-group">
            <label for="baseUrl">API Base URL:</label>
            <input type="text" id="baseUrl" placeholder="https://api.openai.com/v1">
        </div>

        <div class="setting-group">
            <label for="modelSelect">Model:</label>
            <select id="modelSelect">
                <option value="gpt-4">GPT-4</option>
                <option value="gpt-4-turbo-preview">GPT-4 Turbo</option>
                <option value="gpt-3.5-turbo">GPT-3.5 Turbo</option>
                <option value="gpt-3.5-turbo-16k">GPT-3.5 Turbo 16K</option>
                <option value="custom">Custom Model</option>
            </select>
            <input type="text" id="customModel" class="model-input" placeholder="Enter custom model name">
        </div>

        <div class="setting-group">
            <label for="temperature">Temperature:</label>
            <div class="range-container">
                <input type="range" id="temperature" min="0" max="1" step="0.1" value="0.7">
                <span class="value-display" id="temperatureValue">0.7</span>
            </div>
        </div>

        <div class="setting-group">
            <label for="maxTokens">Max Tokens:</label>
            <input type="number" id="maxTokens" min="150" value="1000" step="50">
        </div>

        <button onclick="saveSettings()">Save Settings</button>
        <div id="status" class="status"></div>
    </div>

    <script>
        // Initialize settings
        document.addEventListener('DOMContentLoaded', function() {
            google.script.run
                .withSuccessHandler(loadSettings)
                .getSettings();
        });

        // Handle model select change
        document.getElementById('modelSelect').addEventListener('change', function(e) {
            const customInput = document.getElementById('customModel');
            if (e.target.value === 'custom') {
                customInput.classList.add('show');
            } else {
                customInput.classList.remove('show');
            }
        });

        // Update temperature display
        document.getElementById('temperature').addEventListener('input', function(e) {
            document.getElementById('temperatureValue').textContent = e.target.value;
        });

        function loadSettings(settings) {
            if (!settings) return;
            
            document.getElementById('baseUrl').value = settings.baseUrl || '';
            document.getElementById('temperature').value = settings.temperature || 0.7;
            document.getElementById('temperatureValue').textContent = settings.temperature || 0.7;
            document.getElementById('maxTokens').value = settings.maxTokens || 1000;
            
            const modelSelect = document.getElementById('modelSelect');
            const customModel = document.getElementById('customModel');
            
            if (settings.model) {
                if (['gpt-4', 'gpt-4-turbo-preview', 'gpt-3.5-turbo', 'gpt-3.5-turbo-16k'].includes(settings.model)) {
                    modelSelect.value = settings.model;
                } else {
                    modelSelect.value = 'custom';
                    customModel.value = settings.model;
                    customModel.classList.add('show');
                }
            }
        }

        function saveSettings() {
            const modelSelect = document.getElementById('modelSelect');
            const model = modelSelect.value === 'custom' 
                ? document.getElementById('customModel').value 
                : modelSelect.value;

            const settings = {
                baseUrl: document.getElementById('baseUrl').value,
                model: model,
                temperature: parseFloat(document.getElementById('temperature').value),
                maxTokens: parseInt(document.getElementById('maxTokens').value)
            };

            google.script.run
                .withSuccessHandler(showSuccess)
                .withFailureHandler(showError)
                .saveSettings(settings);
        }

        function showSuccess() {
            const status = document.getElementById('status');
            status.textContent = 'Settings saved successfully!';
            status.className = 'status success';
            status.style.display = 'block';
            setTimeout(() => status.style.display = 'none', 3000);
        }

        function showError(error) {
            const status = document.getElementById('status');
            status.textContent = 'Error saving settings: ' + error;
            status.className = 'status error';
            status.style.display = 'block';
        }
    </script>
</body>
</html>
```

# Project Files Structure
```
gpt-google-sheets-extension/
│
├── appsscript.json          # Configuration file
├── Code.gs                   # Main script file with API calls, menu setup, and core functions
├── Sidebar.html              # Sidebar UI for user prompts and responses
├── Settings.html             # Settings panel UI for API configurations and model parameters
└── README.md                 # Project documentation
```


1. `Code.gs` - Main script file containing:
   - OpenAI API integration
   - Menu creation
   - Core functionality
   - API handling

2. `Sidebar.html` - Single HTML file containing:
   - User interface
   - Styles (in `<style>` tag)
   - Client-side JavaScript (in `<script>` tag)

3. `appsscript.json` - Manifest file (auto-generated):
   - Project configuration
   - Required OAuth scopes
   - Script properties

4. `Settings.html` - Provides the settings interface for updating the API’s 
- base URL
- model
- temperature
- max tokens. 
Contains JavaScript to retrieve and save settings via Code.gs


# Development Setup

1. Create a new Google Apps Script project at script.google.com
2. Create the above files in the project
3. Set up OpenAI API key in Script Properties
4. Deploy as Google Docs add-on
