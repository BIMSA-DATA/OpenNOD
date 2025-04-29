import torch
from transformers import AutoModel, AutoTokenizer
import numpy as np
import pandas as pd
from sklearn.metrics.pairwise import cosine_similarity
from openai import OpenAI
import re

# Configuration
DEVICE = torch.device("cuda:2" if torch.cuda.is_available() else "cpu")

API_KEY = '******' #haihua.xie@gmail.com
clientGPT = OpenAI(
    # defaults to os.environ.get("OPENAI_API_KEY")
    api_key = API_KEY,
)

# Function for Similarity Computation (for Two Excel Files)
def compute_similarity(data1_path, data2_path, model_dir, 
                       output_file="sim-samples.txt", top_k=5):
    df1 = pd.read_excel(data1_path, dtype=str)
    df2 = pd.read_excel(data2_path, dtype=str)
    texts1 = df1['tweet_text'].astype(str).tolist()
    texts2 = df2['tweet_text'].astype(str).tolist()

    # Load the saved model and tokenizer
    encoder = AutoModel.from_pretrained(model_dir).to(DEVICE)
    tokenizer = AutoTokenizer.from_pretrained(model_dir)

    # Precompute embeddings for data2
    with torch.no_grad():
        inputs2 = tokenizer(texts2, padding=True, truncation=True, return_tensors="pt", max_length=128)
        input_ids2, attention_mask2 = inputs2['input_ids'].to(DEVICE), inputs2['attention_mask'].to(DEVICE)
        embeddings2 = encoder(input_ids=input_ids2, attention_mask=attention_mask2).last_hidden_state[:, 0, :].cpu().numpy()

    # Process each query in data1 and find top-k similar indices from data2
    indices_output = []
    processd_num = 0
    for text1 in texts1:
        inputs1 = tokenizer(text1, return_tensors="pt", padding=True, truncation=True, max_length=128).to(DEVICE)
        with torch.no_grad():
            embedding1 = encoder(input_ids=inputs1['input_ids'], attention_mask=inputs1['attention_mask']).last_hidden_state[:, 0, :].cpu().numpy()

        similarities = cosine_similarity(embedding1, embeddings2)[0]
        top_indices = similarities.argsort()[-top_k:][::-1]
        indices_output.append(" ".join(map(str, top_indices)))

        processd_num = processd_num + 1
        if processd_num%100 == 0:
            print(processd_num)

    # Save to file
    with open(output_file, "w") as f:
        for line in indices_output:
            f.write(line + "\n")

def process_tweet(text):
    import re
    # 1) 去除 URL 链接
    text = re.sub(r'http\S+|www\.\S+', '', text)
    # 2) 提取所有 hashtag（包括 # 开头的词）
    hashtags = re.findall(r'#\w+', text)
    # 3) 去掉多余空格
    text = re.sub(r'\s+', ' ', text).strip()
    # 4) 截取前1000个字符（不包括 hashtags）
    truncated_text = text[:330]
    # 5) 确保所有 hashtag 都保留（如果被截掉就加回来）
    for tag in hashtags:
        if tag not in truncated_text:
            truncated_text += f' {tag}'
    return truncated_text.strip()

# Function for Demonstration Selection (for Excel Files)
def opcls_AGI(target_file, label_file, output_file):
    df1 = pd.read_excel(target_file, dtype=str)
    df2 = pd.read_excel(label_file, dtype=str)

    outfile = open(output_file, 'w')

    for index, row in df1.iterrows():
        sim_samples = row['sim_samples'].split(' ')[:3]
        id1, id2, id3 = sim_samples

        text1 = df2[df2['tweet_ID'] == id1]['tweet_text'].str[:1000].values[0] if (df2['tweet_ID'] == id1).any() else None
        text2 = df2[df2['tweet_ID'] == id2]['tweet_text'].str[:1000].values[0] if (df2['tweet_ID'] == id2).any() else None
        text3 = df2[df2['tweet_ID'] == id3]['tweet_text'].str[:1000].values[0] if (df2['tweet_ID'] == id3).any() else None

        stance1 = df2[df2['tweet_ID'] == id1]['stance'].values[0] if (df2['tweet_ID'] == id1).any() else None
        stance2 = df2[df2['tweet_ID'] == id2]['stance'].values[0] if (df2['tweet_ID'] == id2).any() else None
        stance3 = df2[df2['tweet_ID'] == id3]['stance'].values[0] if (df2['tweet_ID'] == id3).any() else None

        prompts = open('prompts_instructions-both-gpt-AGI.txt', encoding='utf-8').read()
        prompts = prompts + (f'Example#1\nTWEET: {text1}\nOUTPUT: {stance1}\n\n'
                             f'Example#2\nTWEET: {text2}\nOUTPUT: {stance2}\n\n'
                             f'Example#3\nTWEET: {text3}\nOUTPUT: {stance3}\n')
        prompts = prompts + ('========== examples ends ==========\n\nTWEET:')
        prompts = prompts + process_tweet(row['tweet_text']) + '\n'

        if index%1000 == 0:
            print(prompts)

        response = clientGPT.chat.completions.create(
            model="gpt-4-turbo-2024-04-09", #$0.1+$0.3/k_token, 128k context
            # model="gpt-4o-mini",  # $0.15+$0.6/M_tokens, 16k context
            # model="gpt-4o",     # $2.5+$10/M_tokens, 128k context
            messages=[
                {
                    "role": "user",
                    "content": prompts
                }
            ],
            temperature=0
        )

        resultstr = response.choices[0].message.content
        if index%1000 == 0:
            print(resultstr)
        outfile.write(resultstr+'\n')

def contains_both_patterns(text):
    pattern = r"\[Hamas:\s*([^\[\];]+)\s*;;\s*([^\[\];]+)\s*\]"
    match = re.search(pattern, text, re.IGNORECASE)

    if match:
        text1, text2 = match.groups()
        if text1.strip().lower() == 'na' or text2.strip().lower() == 'high':
            return True
    return False

def opcls_HIC(target_file, label_file, output_file):
    df1 = pd.read_excel(target_file, dtype=str)
    df2 = pd.read_excel(label_file, dtype=str)
    outfile = open(output_file, 'w')

    prompts = open('prompts_instructions-Hamas-gpt-HIC.txt', encoding='utf-8').read()
    for index, row in df1.iterrows():

        resultstr = ''
        if 1==0: #'hamas: na' not in row['stance_old'].lower():
            resultstr = row['stance_old']
        else:
            print(index)
            print(row['tweet_text'])
            prompts = prompts + process_tweet(row['tweet_text']) + '\n'
            print(prompts)
            # response = clientGPT.chat.completions.create(
            #     model="gpt-4-turbo", #$10+$30/M_token, 128k context
            #     # model="gpt-4o-mini",  # $0.15+$0.6/M_tokens, 16k context
            #     # model="gpt-4o",     # $2.5+$10/M_tokens, 128k context
            #     messages=[
            #         {
            #             "role": "user",
            #             "content": prompts
            #         }
            #     ],
            #     temperature=0
            # )

            # resultstr = response.choices[0].message.content
        outfile.write(resultstr+'\n')

if __name__ == '__main__':
    # compute_similarity("datasets//All_Tweets_list_AGI_PartInfo.xlsx", 
    #                    "datasets//labeled-AGI-600.xlsx", 
    #                    "saved_model//AGI","sim-samples-AGI.txt",5)

    #opcls_AGI("datasets//All_Tweets_list_AGI_PartInfo.xlsx", "datasets//labeled-AGI-600.xlsx","AGI.txt")
    opcls_HIC("datasets//All_Tweets_list_HIC_PartInfo1.xlsx", "datasets//labeled-HIC-Hamas-600.xlsx", "HIC-gpt4turbo-39545.txt")
