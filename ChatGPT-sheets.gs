
// main function
function GPT(text, prompt, systemPrompt='', maxTokens=200, temperature=0.0) {

  // EITHER get the API key from the "API-key" sheet
  const apiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("API-key");
  const apiKey = apiSheet.getRange(3,1).getValue() // Cell A3
  // OR set API key here in the script
  //const apiKey = "..."
  
  // set default system prompt
  if (systemPrompt === '') {
    systemPrompt = "You are a helpful, factual, rational research assistant. You do not respond anything other than the exact response to your task."
  }
  Logger.log("System prompt: " + systemPrompt)
  // reset default values in case user puts empty string/cell as argument
  if (maxTokens === '') {
    maxTokens = 200
  }
  Logger.log("maxTokens: " + maxTokens)
    if (temperature === '') {
    temperature = 0.0
  }
  Logger.log("temperature: " + temperature)

  // configure the API request to OpenAI
  const data = {
    "model": "gpt-3.5-turbo",
    "messages": [
      {"role": "system", "content": systemPrompt},
      {"role": "user", "content": "Here is a text:\n" + text + "\n" + prompt}
    ],
    "temperature": temperature,
    "max_tokens": maxTokens,
    "n": 1,
    "top_p": 1.0,
    "frequency_penalty": 0.0,
    "presence_penalty": 0.0,
  };
  const options = {
    'method' : 'post',
    'contentType': 'application/json',
    'payload' : JSON.stringify(data),
    'headers': {
      Authorization: 'Bearer ' + apiKey,
    },
  };
  const response = UrlFetchApp.fetch(
    'https://api.openai.com/v1/chat/completions',
    options
  );
  Logger.log(response.getContentText());

  // Send the API request 
  const result = JSON.parse(response.getContentText())['choices'][0]['message']['content'];
  Logger.log(result)


  // use OpenAI moderation API to block problematic outputs.
  // see: https://platform.openai.com/docs/guides/moderation/overview
  var optionsModeration = {
    'payload' : JSON.stringify({"input": result}),
    'method' : 'post',
    'contentType': 'application/json',
    'headers': {
      Authorization: 'Bearer ' + apiKey,
    },
  }
  const responseModeration = UrlFetchApp.fetch(
    'https://api.openai.com/v1/moderations',
    optionsModeration
  );
  Logger.log(responseModeration)
  // Send the API request  for moderation API
  const resultFlagged = JSON.parse(responseModeration.getContentText())['results'][0]['flagged'];
  Logger.log(resultFlagged)
  // do not return result, if moderation API determined that the result is problematic
  if (resultFlagged === true) {
    return "The OpenAI moderation API blocked the response."
  }

  // try parsing the output as JSON. 
  // If JSON, then JSON values are returned across the cells to the right of the selected cell. 
  // If not JSON, then the full result is directly returned
  try {
    const resultJson = JSON.parse(result);
    Logger.log(resultJson);
    // Initialize an empty array to hold the values
    const resultJsonValues = [];
    // Loop through the object using for...in loop
    for (const key in resultJson) {
      // Push the value of each key into the array
      resultJsonValues.push(resultJson[key]);
      Logger.log(key);
      Logger.log(resultJson[key])
    }
    console.log("Final result: " + resultJsonValues)
    return [resultJsonValues]
  // Just return full response if result not parsable as JSON
  } catch (e) {
    Logger.log(e);
    console.log("Final result: ", result);
    return result;
  }

}

// Add ChatGPT Menu
// commenting this out to remove permission requirements for running the function
/*const onOpen = () => {
const ui = SpreadsheetApp.getUi();
ui.createMenu("ChatGPT")
  .addItem("This notebook was created by Moritz Laurer. Click here to see the source-code on GitHub.", "openUrl")
  // This script was inspired by the great blog post by Sarah Tamsin: https://sarahtamsin.com/integrate-chatgpt-with-google-sheets/
  .addToUi();
};

const openUrl = () => {
  const url = "https://github.com/MoritzLaurer/ChatGPT-google-sheets";
  const html = "<script>window.open('" + url + "', '_blank');google.script.host.close();</script>";
  const ui = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(ui, "Redirecting...");
};
*/


