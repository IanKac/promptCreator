This application is designed for:
1. Creating prompts based on provided template,
2. Sending said prompts to chatGPT API and reciving data,
3. Exporting recived data to WORD document.

To create prompts you need to create prompt in 'Prompts' in cell A2, with marked places for variables in square brackets ('[',']'). 
Variables need to match exactly with column names in 'Data' tab. Data can be whatever you want but need to start in B column.
Created prompts will apear in 'Prompts' tab in column B.
Requesting data takes some time but when it is done it will apear in 'Prompts' tab in C column.

All logs, prompts and extracts are in respective folders. Those folders are created automatically on first go in current location of this application therefore I suggest putting it in folder to avoid mess.
API needs to be created in location '..\API\' and be named 'ChatGPT_APIkey.txt' for application to find it.
