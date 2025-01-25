// this is a google apps script
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Setup OpenAI')
      .addItem('Set Secret Key', 'set_openai_secret_key')
    .addToUi();
}

function set_openai_secret_key() {
    // check if this google sheet has a secret property named "openai_secret_key"
    var secret_key = PropertiesService.getScriptProperties().getProperty("openai_secret_key");
    var additional_info = ""

    if (secret_key) {
        let first_5_and_last_5_chars = secret_key.substring(0, 5) + "..." + secret_key.substring(secret_key.length - 5);
        additional_info = `OpenAI secret key already set: 
        
        ${first_5_and_last_5_chars} 
        
        If you'd like to change it then enter a new one below`;
    }

    let info_message = 
        "=== OpenAI Secret Key Setup ===\n\n" +
        "To get your OpenAI secret key:\n" +
        "1. Visit: https://platform.openai.com/api-keys\n" +
        "2. Create a new secret key\n" +
        "3. Copy and paste it below\n\n" +
        (additional_info ? additional_info + "\n\n" : "")

    // first show a message to user to enter the secret key
    var ui = SpreadsheetApp.getUi();

    // then get the secret key from the user
    var secret_key = ui.prompt(info_message).getResponseText();
    
    if (secret_key.length < 10) {
        ui.alert("Please enter a valid OpenAI secret key");
        return;
    } else {
        try {
            PropertiesService.getScriptProperties().setProperty("openai_secret_key", secret_key);
            SpreadsheetApp.getUi().alert('Success! OpenAI secret key has been saved.');
        } catch (error) {
            SpreadsheetApp.getUi().alert('Error: Failed to save OpenAI secret key. Please try again.');
        }
    }
}

/**
 * Makes an API call to OpenAI's GPT model and returns the response
 * @param {string} input - The text prompt to send to GPT
 * @return {string} The GPT response
 * @customfunction
 */
function SheetGPT(input) {
  console.log("SheetGPT called with input:", input);
  
  // Input validation
  if (!input || (typeof input === 'string' && input.trim() === '') || 
      (Array.isArray(input) && input.flat().join('').trim() === '')) {
    console.log("Empty input received, skipping API call");
    return "#NOINPUT";
  }
  
  // Get the API key
  const apiKey = PropertiesService.getScriptProperties().getProperty("openai_secret_key");
  if (!apiKey) {
    console.error("API key not found in DocumentProperties");
    return "Error: OpenAI API key not set. Please use 'Setup OpenAI' menu to set your API key.";
  }
  console.log("API key found, making request to OpenAI");

  try {
    // Handle input if it's an array (multiple cells selected)
    let processedInput = input;
    if (Array.isArray(input)) {
      processedInput = input.flat().join(" ");
    }
    console.log("Processed input:", processedInput);

    const payload = {
      'model': 'gpt-4o-mini',
      'messages': [
        {
          'role': 'system',
          'content': "You are a helpful assistant processing data from Google Sheets. Process the input and provide a concise response."
        },
        {
          'role': 'user',
          'content': processedInput
        }
      ],
      'temperature': 0.7,
      'max_tokens': 150
    };
    
    console.log("Request payload:", JSON.stringify(payload));
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true  // Added to get full error messages
    });
    
    console.log("Response status:", response.getResponseCode());
    
    const json = JSON.parse(response.getContentText());
    if (json.error) {
      console.error("API Error:", json.error);
      return "Error: " + json.error.message;
    }
    
    const result = json.choices[0].message.content.trim();
    console.log("Successfully processed response:", result);
    return result;
  } catch (error) {
    console.error("Error in SheetGPT:", error.toString());
    console.error("Full error object:", JSON.stringify(error));
    return "Error: " + error.toString();
  }
}

