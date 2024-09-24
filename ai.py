import os
import pandas as pd
from openai import AzureOpenAI
from dotenv import load_dotenv

# Load the .env file
load_dotenv()

client = AzureOpenAI(
    azure_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT"), 
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),  
    api_version="2024-05-01-preview"
)

def chat_with_openai(user_input):
    response = client.chat.completions.create(
        model=os.getenv("AZURE_OPENAI_DEPLOYMENT_ID"),
        messages=[
            {"role": "system", "content": """You are a Business Central AL Developer. You will be provided with code snippets. 
             Based on the functionality within the code, respond with the purpose of the modification related to the food and beverage industry. Response should be short and clear. Just reason title is fine.
              Do not explain or provide technical details—only state the purpose of the modification. If there is any commented code, try to understand the functionality and provide the purpose. 
             If you are unsure about the purpose, just respond with 'I am not sure'."""},
            {"role": "user", "content": user_input}
        ]
    )
    return response.choices[0].message.content

def open_excel(file_path):
    df = pd.read_excel(file_path)
    second_column_name = df.columns[1]

    for index, row in df.iterrows():
        text = row[second_column_name]

        response = chat_with_openai(text)
        print(response)

        df.loc[index, 'Response'] = response

    # Remove rows where the 'Response' column contains "I am not sure."
    # df = df[~df['Response'].str.contains("I am not sure.")]

    df.to_excel(file_path, index=False)

open_excel('./baker.xlsx')