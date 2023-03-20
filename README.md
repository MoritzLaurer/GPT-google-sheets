# Use GPT4, ChatGPT and other LLMs in Google Sheets 🤖🧑 

This repository provides custom functions for Google Sheets that enables you to use Large Language Models (LLMs) like ChatGPT or GPT4 without any coding knowledge. With these functions, you can turn large amounts of unstructured text into structured data in seconds. 


## Update 20.03.23

**The repository now supports the following models and functions:**
* `=GPT4()`. This function uses gpt-4.
* `=CHATGPT()`. This function uses gpt-3.5-turbo (ChatGPT). 
* The following functions will be supported soon: `=PALM()` using [Google's PaLM model](https://blog.google/technology/ai/ai-developers-google-cloud-workspace/) and `=CLAUDE()` using [Anthropic's Claude model](https://console.anthropic.com/docs).

**The functions support three main use-cases:**
* Produce **unstructured output** in one cell. For this use-case use a standard prompt. E.g.: `=CHATGPT(<input_text>, "Summarize the text", "You are a helpful assistant")`. The prompt will simply output a summary into one cell. 
* Produce **structured horizontal output** in one row. For this use-case, use a JSON prompt. E.g.: `=CHATGPT(<input_text>, "Is the text positive, negative or neutral? First, reason step-by-step for your response. Second, provide one of these responses "Positive", "Negative", "Neutral". ALWAYS (!) respond in the following JSON format: {"reason": "...", "sentiment": "..."}, "You are a JSON producing text analyser and only output JSON.")`. This prompt will output the `reason` into the first cell and the `sentiment` into the second cell to the right. 
* Produce **structured table output** in both rows and columns. For this use-case, use a markdown table prompt. E.g.: `=CHATGPT(<input_text>, "Convert the text into a table.", "You are a table producing text analyser and only output markdown tables.")`. This prompt will output a clean table into Google Sheet columns and rows. 

**Helper functions:**
* `=GPTCOST()` returns a rough cost estimate for each function call. For example `=GPTCOST("gpt4", 100, <long_text>, <prompt>, <system_prompt>)` will estimate the cost for maximum 100 output tokens and tokens from a long_text to analyse, a prompt and a system_prompt.  


## Example video for 15.03.23 version

https://user-images.githubusercontent.com/41862082/225295311-d851c1a1-2a3a-411b-86c7-f511de629880.mp4

Note that the `=GPT()` function is depricated. Use `=CHATGPT()` or `=GPT4()` instead.

## Example sheet and usage instructions
Example use-cases with instructions and example data are available in this Google Sheet: 
https://docs.google.com/spreadsheets/d/1_om1a0wv6boajKroaAlupayKPM32Bm1M7HjZ98XcDRA/edit?usp=sharing

You can use the function in two main ways: 
1. You can create a copy of the sheet above and directly run the function on the example data. Open the link and in the menu bar click > File > Make a copy. You can then run the function in your own copy of the sheet. 
2. You can copy the `GPT.gs` script from this repository and paste it in any of your own Google Sheets. In the menu bar of your own sheet, click Extensions > Apps Script > Add a new file and paste the script. 

Note that you need to add your own OpenAI API key to a sheet called 'API-key' in cell A3. If you use the function in your own scripts, create a new sheet called "API-key" and copy your API key in cell A3 (see the example script). If you do not have an OpenAI API key, you need to create an account and a new API key [here](https://platform.openai.com/account/api-keys). For details on pricing, see [here](https://openai.com/pricing).

More detailed explanations for how to use the functions is provided in the 'examples' sheet. 

This repository was created by Moritz Laurer. If it is useful for you, give it a 🌟. If you have questions, reach out on [LinkedIn](https://www.linkedin.com/in/moritz-laurer/) or [Twitter](https://twitter.com/MoritzLaurer).


## License
The script is licensed under the permissive [Responsible AI End-User License v0.1 (RAIL)](https://www.licenses.ai/ai-licenses). You are free to use it for pretty much any non-commercial or commercial purpose, but it cannot be used for criminal or similar activities. Note that I have added a call to the OpenAI moderation API to prevent misuse of the model. 
