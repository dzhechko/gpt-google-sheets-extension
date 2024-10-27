// Store your API key securely in Script Properties
function setApiKey(apiKey) {
  if (!apiKey) {
    throw new Error('API key cannot be empty');
  }
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
}

// Get API key from Script Properties
function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('API key not found. Please set your OpenAI API key in settings.');
  }
  return apiKey;
}

// Get settings from Script Properties
function getSettings() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const settings = scriptProperties.getProperty('API_SETTINGS');
  
  if (!settings) {
    // Return default settings if none are saved
    return {
      baseUrl: 'https://api.openai.com/v1',
      model: 'gpt-3.5-turbo',
      temperature: 0.7,
      maxTokens: 1000
    };
  }
  
  return JSON.parse(settings);
}

// Save settings to Script Properties
function saveSettings(settings) {
  if (!settings) {
    throw new Error('Settings object cannot be empty');
  }
  
  // Validate settings
  if (!settings.baseUrl || !settings.model) {
    throw new Error('Base URL and model are required');
  }
  
  if (settings.temperature < 0 || settings.temperature > 1) {
    throw new Error('Temperature must be between 0 and 1');
  }
  
  if (settings.maxTokens < 150) {
    throw new Error('Max tokens must be at least 150');
  }
  
  PropertiesService.getScriptProperties().setProperty('API_SETTINGS', JSON.stringify(settings));
}

// Function to call OpenAI API
function callOpenAI(prompt) {
  if (!prompt) {
    throw new Error('Prompt cannot be empty');
  }

  const apiKey = getApiKey();
  const settings = getSettings();
  const apiUrl = `${settings.baseUrl}/chat/completions`;
  
  const requestBody = {
    model: settings.model,
    messages: [
      { role: 'user', content: prompt }
    ],
    temperature: settings.temperature,
    max_tokens: settings.maxTokens
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
    throw new Error(`Failed to call OpenAI API: ${error.message}`);
  }
}

// Test function to verify API connection
function testOpenAI() {
  try {
    const result = callOpenAI('What is the capital of France?');
    Logger.log(result);
    return result;
  } catch (error) {
    Logger.log(`Test failed: ${error.message}`);
    throw error;
  }
}

// Create menu when the spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GPT Помощник')
    .addItem('Показать панель ассистента', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Инструменты')
      .addItem('Анализ выбранных ячеек', 'analyzeSelection')
      .addItem('Создать сводку', 'generateSummary'))
    .addSeparator()
    .addItem('Настройки', 'showSettings')
    .addToUi();
}

// Function to show sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('GPT Ассистент')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Function to analyze selected cells
function analyzeSelection() {
  // Get the currently selected range in the spreadsheet
  const selection = SpreadsheetApp.getActiveRange();
  
  // Check if any cells are selected
  if (!selection) {
    SpreadsheetApp.getUi().alert('Пожалуйста, сначала выберите ячейки.');
    return;
  }
  
  try {
    // Get the values from selected cells as a 2D array
    const values = selection.getValues();
    const numRows = values.length;
    const numCols = values[0].length;
    
    const prompt = `Analyze this spreadsheet data (${numRows} rows × ${numCols} columns):
    ${JSON.stringify(values)}
    
    Please provide a detailed analysis including:
    1. Data Structure Overview:
       - Type of data in each column
       - Data range and distribution
    
    2. Key Patterns and Trends:
       - Main trends in the data
       - Any correlations between columns
       - Unusual or outlier values
    
    3. Statistical Insights:
       - Key statistics where relevant (averages, min/max, etc.)
       - Distribution patterns
    
    4. Notable Observations:
       - Any interesting findings
       - Potential data quality issues
       - Anomalies or unexpected patterns

    Please format the response in a clear, structured way.`;
    
    // Call OpenAI API with the prompt
    const analysis = callOpenAI(prompt);
    
    // Display results in a modal dialog
    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h3>Результаты анализа</h3>
        <div style="white-space: pre-wrap; background-color: #f8f9fa; padding: 12px; border-radius: 4px;">
          ${analysis}
        </div>
      </div>
    `)
    .setWidth(600)
    .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Результаты анализа');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
    Logger.log('Analysis error: ' + error.message);
  }
}

// Function to generate summary of the entire sheet
function generateSummary() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  
  try {
    const values = dataRange.getValues();
    const headers = values[0];
    const numRows = values.length;
    const numCols = headers.length;
    
    const prompt = `Generate a summary of this spreadsheet:
    Sheet name: ${sheet.getName()}
    Headers: ${headers.join(', ')}
    Number of rows: ${numRows}
    Number of columns: ${numCols}
    
    Data sample (first few rows): ${JSON.stringify(values.slice(0, 3))}
    
    Please provide:
    1. A brief overview of the data structure
    2. Key insights or patterns
    3. Any recommendations for analysis`;
    
    const summary = callOpenAI(prompt);
    
    // Show results in a modal dialog with formatted HTML
    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h3>Sheet Summary</h3>
        <div style="white-space: pre-wrap; background-color: #f8f9fa; padding: 12px; border-radius: 4px;">
          ${summary}
        </div>
      </div>
    `)
    .setWidth(600)
    .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Sheet Summary');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    Logger.log('Summary error: ' + error.message);
  }
}

// Function to show settings dialog
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройки GPT Помощника');
}
