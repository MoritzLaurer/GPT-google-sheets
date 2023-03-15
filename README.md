# Use ChatGPT in Google Sheets without any code 🤖🧑 

This repository provides a custom function for Google Sheets '=GPT()' that enables you to use ChatGPT without any coding knowledge. With this function, you can turn large amounts of unstructured text into structured data in seconds. 

https://user-images.githubusercontent.com/41862082/225295311-d851c1a1-2a3a-411b-86c7-f511de629880.mp4


## Example sheet and instructions
Example use-cases with instructions and example data are available in this Google Sheet: 
https://docs.google.com/spreadsheets/d/1_om1a0wv6boajKroaAlupayKPM32Bm1M7HjZ98XcDRA/edit?usp=sharing

You can use the function in two main ways: 
1. You can create a copy of the sheet above and directly run the function on the example data. Open the link and in the menu bar click > File > Make a copy. You can then run the function in your own copy of the sheet. 
2. You can copy the `ChatGPT-sheets.gs` script from this repository and paste it in any of your own Google Sheets. In the menu bar of your own sheet, click Extensions > Apps Script > Add a new file and paste the script. 

Note that you need to add your own OpenAI API key to the 'API-key' sheet. If you use the function in your own scripts, create a new sheet called "API-key" and copy your API key in cell A3 (see the example script). If you do not have an OpenAI API key, you need to create an account and a new API key [here](https://platform.openai.com/account/api-keys). The script uses the gpt-3.5-turbo model (ChatGPT), which costs 0.002 per 1000 tokens. For details on pricing, see [here](https://openai.com/pricing).

More detailed explanations for how to use the '=GPT()' function is provided in the 'examples' sheet. 

This repository was created by Moritz Laurer. If it is useful for you, give it a 🌟. If you have questions, reach out on [LinkedIn](https://www.linkedin.com/in/moritz-laurer/) or [Twitter](https://twitter.com/MoritzLaurer).


## License
The script is licensed under the permissive ["Unlicense" license](https://github.com/MoritzLaurer/ChatGPT-google-sheets/blob/master/LICENSE), so you are free to use it for pretty much any non-commercial or commercial purpose. Note that I have added a call to the OpenAI moderation API to prevent misuse of the model. 
