import pandas as pd
import re
import json
from openai import OpenAI
from qianfan import Qianfan
from llamaapi import LlamaAPI

dataDir = 'D:\\MyPythonScripts\\TweetsStanceAnalysis\\datasets\\'

# Baidu Client
client_DS = Qianfan(
    api_key="******"
    # app_id="" # optional. If not filled in, the default appid will be used.
)

# GPT Client
API_KEY = '******' #haihua.xie@gmail.com
client_GPT = OpenAI(
    # defaults to os.environ.get("OPENAI_API_KEY")
    api_key = API_KEY,
)

# Llama Client
API_TOKEN = '******'
client_LA = LlamaAPI(API_TOKEN)

# Function to remove URLs from a text string
def remove_urls(text):
    if not isinstance(text, str):  # Ensure the input is a string
        text = str(text)
    text = re.sub(r'http\S+|www\S+', '', text).strip()  # Remove URLs
    return text[:1000]  # Limit text to 1000 characters

def paraphraseUsingGPT(inputText):
    prompts = ('For the given Input Text, generate a rephrased version that preserves its original '
               'semantic meaning, sentiment polarity, and sentiment intensity. Directly output the '
               'rephrased version without adding any explanatory content. \n Input Text: ')
    prompts = prompts + inputText
    print(prompts + '\n =======  ======= ======= ======= =======\n')

    response = client_GPT.chat.completions.create(
        model="gpt-4o-mini",  # $0.15+$0.6/M_tokens, 16k context
        # model="gpt-4o",     # $2.5+$10/M_tokens, 128k context
        messages=[
            {
                "role": "user",
                "content": prompts
            }
        ],
        temperature=1.0
    )

    resultstr = response.choices[0].message.content
    #print(resultstr)
    print('\n =======  ======= ======= ======= =======\n')
    return resultstr

def paraphraseUsingDS(inputText):
    prompts = ('For the given Input Text, generate a rephrased version that preserves its original '
               'semantic meaning, sentiment polarity, and sentiment intensity. Directly output the '
               'rephrased version without adding any explanatory content. \n Input Text: ')
    prompts = prompts + inputText
    print(prompts + '\n =======  ======= ======= ======= =======\n')

    chat_completion = client_DS.chat.completions.create(
        #model="ernie-4.0-turbo-8k-latest",
        model="deepseek-v3",
            messages=[
            {
                "role": "user",
                "content": prompts
            }
        ],
        temperature=1.0
    )

    resultstr = chat_completion.choices[0].message.content
    #print(resultstr)
    print('\n =======  ======= ======= ======= =======\n')
    return resultstr

def paraphraseUsingLA(inputText):
    prompts = ('For the given Input Text, generate a rephrased version that preserves its original '
               'semantic meaning, sentiment polarity, and sentiment intensity. Directly output the '
               'rephrased version without adding any explanatory content. \n Input Text: ')
    prompts = prompts + inputText
    print(prompts + '\n =======  ======= ======= ======= =======\n')

    # Build the API request
    api_request_json = {
        # "model": "llama3.1-70b",
        "model": "llama3.1-405b",

        "messages": [
            {"role": "user", "content": prompts},
        ],

        "temperature": 1.0,
    }

    # Execute the Request
    response = client_LA.run(api_request_json)
    print('\n =======  ======= ======= ======= =======\n')
    return response.json()["choices"][0]["message"]["content"]

def paraphraseTweets(input_file, output_file, GPT_DS_LA='LA'):
    # Read the Excel file, ensuring all data is treated as text
    df = pd.read_excel(input_file, dtype=str)

    # Create a new DataFrame to store expanded rows
    expanded_rows = []

    # Process each row individually
    for index, row in df.iterrows():
        cleaned_text = remove_urls(row['tweet_text'])
        for i in range(1, 4):  # Create 3 versions with numbered prefixes
            new_row = row.copy()
            if GPT_DS_LA == 'GPT':
                parapTweet = paraphraseUsingGPT(cleaned_text)
            elif GPT_DS_LA == 'DS':
                parapTweet = paraphraseUsingDS(cleaned_text)
            else:
                parapTweet = paraphraseUsingLA(cleaned_text)
            new_row['tweet_text'] = f"{parapTweet}"
            expanded_rows.append(new_row)

    # Convert list to DataFrame
    expanded_df = pd.DataFrame(expanded_rows)

    # Save to a new Excel file, ensuring text format
    expanded_df.to_excel(output_file, index=False, engine='openpyxl')

    print("Processing complete. Cleaned and expanded data saved.")

def merge_xlsx(xlsx_file1, xlsx_file2, output_file):
    df1 = pd.read_excel(xlsx_file1, dtype=str)
    df2 = pd.read_excel(xlsx_file2, dtype=str)

    combined_df = pd.concat([df1, df2])

    # Remove duplicates based on tweet_text, and
    # keep the row corresponding to the first occurrence of tweet_text.
    unique_df = combined_df.drop_duplicates(subset=['tweet_text'], keep='first')

    unique_df.to_excel(output_file, index=False)

if __name__ == '__main__':
    # paraphraseTweets(dataDir + "new-HIC-Hamas-600.xlsx",
    #                  dataDir + "para-ds-new-HIC-Hamas-600.xlsx", 'ds')

    merge_xlsx(dataDir + "para-new-HIC-Israel-600-1.xlsx",
               dataDir + "para-la-new-HIC-Israel-600.xlsx",
               dataDir + "para-new-HIC-Israel-600.xlsx")

