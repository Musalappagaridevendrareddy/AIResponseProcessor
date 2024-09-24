# AIResponseProcessor

This Python script is used to fetch specific code patterns from files in a specified directory and interact with Azure OpenAI.

## Script: ai.py

The `ai.py` script contains the main functionality of the project. It uses the Azure OpenAI API to chat with an AI model and processes an Excel file based on the AI's responses.

## Usage

To use this script, simply run the `ai.py` file. The script will process an Excel file named 'baker.xlsx' in the same directory, sending each text from the second column to the Azure OpenAI API, adding the response to the end of the row, and removing rows where the 'Response' column contains "I am not sure.".

## Requirements

This script requires the following Python libraries:

- os
- pandas
- openai
- dotenv

You also need to have a `.env` file in the same directory as the script, containing your Azure OpenAI endpoint, API key, and deployment ID.

## Output

The script outputs an updated 'baker.xlsx' Excel file with the AI's responses in the 'Response' column.
