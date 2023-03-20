
// functions to call specifically on ChatGPT using the OpenAIGPT base function
function CHATGPT(text, prompt, systemPrompt='', maxTokens=200, temperature=0.0, model='gpt-3.5-turbo') {
  console.log("Executing function: ChatGPT")
  return OpenAIGPT(text, prompt, systemPrompt=systemPrompt, maxTokens=maxTokens, temperature=temperature, model=model)
}
// functions to call specifically on GPT-4 using the OpenAIGPT base function
function GPT4(text, prompt, systemPrompt='', maxTokens=200, temperature=0.0, model='gpt-4') {
  console.log("Executing function: GPT4")
  return OpenAIGPT(text, prompt, systemPrompt=systemPrompt, maxTokens=maxTokens, temperature=temperature, model=model)
}

// base function for OpenAI GPT models with chat format
// this currently is 'gpt-3.5-turbo' (=ChatGPT) and GPT-4  (as of 20.03.2023) 
function OpenAIGPT(text, prompt, systemPrompt='', maxTokens=200, temperature=0.0, model='gpt-3.5-turbo') {

  // EITHER get the API key from the "API-key" sheet
  const apiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("API-key");
  const apiKey = apiSheet.getRange(3,1).getValue() // Cell A3
  // OR set API key here in the script
  //const apiKey = "..."
  
  // set default system prompt
  if (systemPrompt === '') {
    systemPrompt = "You are a helpful assistant."
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
    "model": model, // "gpt-3.5-turbo",
    "messages": [
      {"role": "system", "content": systemPrompt},
      {"role": "user", "content": text + "\n" + prompt}
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
  Logger.log("Model API response: " + response.getContentText());

  // Send the API request 
  const result = JSON.parse(response.getContentText())['choices'][0]['message']['content'];
  Logger.log("Content message response: " + result)


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
  Logger.log("Moderation API response: " + responseModeration)
  // Send the API request  for moderation API
  const resultFlagged = JSON.parse(responseModeration.getContentText())['results'][0]['flagged'];
  Logger.log('Output text flagged? ' + resultFlagged)
  // do not return result, if moderation API determined that the result is problematic
  if (resultFlagged === true) {
    return "The OpenAI moderation API blocked the response."
  }

  // try parsing the output as JSON or markdown table. 
  // If JSON, then JSON values are returned across the cells to the right of the selected cell.
  // If markdown table, then the table is spread across the cells horizontally and vertically
  // If not JSON or markdown table, then the full result is directly returned in a single cell
  
  // JSON attempt
  if (result.includes('{') && result.includes('}')) {
    // remove boilerplate language before and after JSON/{} in case model outputs characters outside of {} delimiters
    // assumes that only one JSON object is returned and it only has one level and is not nested with multiple levels (i.e. not multiple {}{} or {{}})
    let cleanedJsonString = '{' + result.split("{")[1].trim()
    cleanedJsonString = cleanedJsonString.split("}")[0].trim() + '}'
    Logger.log(cleanedJsonString)
    // try parsing
    try {
      // parse JSON
      const resultJson = JSON.parse(cleanedJsonString);
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
      console.log("Final JSON output: " + resultJsonValues)
      return [resultJsonValues]
    // Just return full response if result not parsable as JSON
    } catch (e) {
      Logger.log("JSON parsing did not work. Error: " + e);
    }
  }

  // .md table attempt
  // determine if table by counting "|" characters. 6 or more
  if (result.split("|").length - 1 >= 6) {
    let arrayNested = result.split('\n').map(row => row.split('|').map(cell => cell.trim()));
    // remove the first and last element of each row, they are always empty strings due to splitting
    // this also removes boilerplate language before and after the table
    arrayNested = arrayNested.map((array) => array.slice(1, -1));
    // remove nested subArray if it is entirely empty. This happens when boilerplate text is separated by a "\n" from the rest of the text
    arrayNested = arrayNested.filter((subArr) => subArr.length > 0);
    // remove the "---" header indicator
    arrayNested = arrayNested.filter((subArr) => true != (subArr[0].includes("---") && subArr[1].includes("---")))
    console.log("Final output as array: ", arrayNested)
    return arrayNested
  }

  // if neither JSON nor .md table
  console.log("Final output without JSON or .md table: ", result)
  return result; 

}




// function for rough cost estimates per request
function GPTCOST(model='', maxOutputTokens=100, text='', prompt='', systemPrompt='') {
/*function GPTCOST() {
  let text = "Here is some test text. Here is some test text.Here is some test text.Here is some test text.Here is some test text.Here is some test text.Here is some test text."
  let prompt = "This is some prompt.Here is some test text.Here is some test text.Here is some test text.Here is some test text."
  let systemPrompt = ""
  let maxOutputTokens = 200
  let model = "gpt-4-32k" //"gpt-3.5-turbo"*/

  let inputAll = `${text} ${prompt} ${systemPrompt}`
  console.log('Full input string:\n', inputAll)
  let inputCharacters = inputAll.length
  console.log('Characters:\n', inputCharacters)

  // OpenAI
  if (model.toLowerCase().includes("gpt")) {
    // simple approximation where 1 token === 4 characters (because cannot install tiktoken in Apps Script) https://platform.openai.com/docs/introduction/tokens
    let nInputTokens = Math.round(inputCharacters / 4)
    // based on the OpenAI cookbook, boilerplate 4+2 tokens needed to be added to each message: https://github.com/openai/openai-cookbook/blob/main/examples/How_to_count_tokens_with_tiktoken.ipynb
    nInputTokens += 6
    console.log("N tokens estimated:\n", nInputTokens)
    // cost estimate in USD  https://openai.com/pricing
    // ChatGPT
    if (model.toLowerCase().includes("gpt-3.5-turbo") || model.toLowerCase().includes("chatgpt")) {
      let costPerToken = 0.002 / 1000
      let costForRequest = costPerToken * (nInputTokens + maxOutputTokens)
      costForRequest = +costForRequest.toFixed(4)  // round to 4 decimals
      console.log("Cost for request in USD:\n", costForRequest)
      return costForRequest
    // gpt-4 model references: https://platform.openai.com/docs/models/gpt-4
    } else if (model.toLowerCase().includes("gpt-4-32k")) {
      // for long text 32k token version
      let costPerInputToken = 0.06 / 1000
      let costPerOutputToken = 0.16 / 1000
      let costForRequest = (costPerInputToken * nInputTokens) + (costPerOutputToken * maxOutputTokens)
      costForRequest = +costForRequest.toFixed(4)  // round to 4 decimals
      console.log("Cost for request in USD:\n", costForRequest)
      return costForRequest
    } else if (model.toLowerCase().includes("gpt-4") || model.toLowerCase().includes("gpt4")) {
      // for standard 8k token version
      let costPerInputToken = 0.03 / 1000
      let costPerOutputToken = 0.06 / 1000
      let costForRequest = (costPerInputToken * nInputTokens) + (costPerOutputToken * maxOutputTokens)
      costForRequest = +costForRequest.toFixed(4)  // round to 4 decimals
      console.log("Cost for request in USD:\n", costForRequest)
      return costForRequest
    } else {
      console.log(`Model name ${model} is not supported. Choose one from: gpt-3.5-turbo, gpt-4, gpt-4-32k from the official OpenAI documentation`)
      return `Model name ${model} is not supported. Choose one from: gpt-3.5-turbo, gpt-4, gpt-4-32k from the official OpenAI documentation`
    }
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


