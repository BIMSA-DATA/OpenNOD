import os
import pandas as pd
from datetime import datetime, timedelta
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import torch
from sentence_transformers.util import cos_sim

# To be run on a GPU server.
# Set the GPU number to be used
os.environ["CUDA_VISIBLE_DEVICES"] = "1"

# Check if there is an available GPU
device = torch.device("cuda" if torch.cuda.is_available() else "cpu")

# Load the SBERT model and move it to the specified GPU.
model = SentenceTransformer('all-mpnet-base-v2', device=device)
#model = SentenceTransformer('paraphrase-multilingual-mpnet-base-v2', device=device)
#model = SentenceTransformer('distiluse-base-multilingual-cased-v2', device=device)

def read_excel_file(file_path):
    """
    Read an Excel file and read all columns in string format.
    :param file_path: The path of the Excel file
    :return: A DataFrame containing tweet data
    """
    df = pd.read_excel(file_path, dtype=str)
    df['tweet_pub_time'] = pd.to_datetime(df['tweet_pub_time'])
    return df

def read_news_files(folder_path):
    """
    Read all TXT files in the specified folder.
    :param folder_path: Path of the folder
    :return: A dictionary containing news data, where the key is the date and the value is the list of news for that day.
    """
    news_dict = {}
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.txt'):
            date_str = file_name.replace('.txt', '')
            date = datetime.strptime(date_str, '%Y-%m-%d')
            with open(os.path.join(folder_path, file_name), 'r', encoding='utf-8') as file:
                news = []
                for line in file.readlines():
                    if len(line.strip()) >= 5:
                        news.append(line.strip())
            news_dict[date] = news
    return news_dict

def mark_tweets(df, news_dict):
    """
    Label tweets
    :param df: A DataFrame containing tweet data
    :param news_dict: A dictionary containing news data
    :return: The labeled DataFrame
    """
    df['mark_obj'] = 0

    for news_date, news_list in news_dict.items():
        start_date = news_date - timedelta(days=1)
        end_date = news_date + timedelta(days=4)

        relevant_tweets = df[(df['tweet_pub_time'] >= start_date) &
                             (df['tweet_pub_time'] <= end_date) &
                             (df['stance_new'].str.contains('objective'))]

        if relevant_tweets.empty:
            continue

        tweet_texts = relevant_tweets['tweet_text'].tolist()
        tweet_indices = relevant_tweets.index.tolist()

        # 💡 Encode tweets once
        tweet_embeddings = model.encode(tweet_texts, batch_size=128, convert_to_tensor=True, device=device)

        for news in news_list:
            if len(news) < 5:
                continue

            print(f"[News] {news}")

            # Encode news once
            news_embedding = model.encode(news, convert_to_tensor=True, device=device)

            # Cosine similarity in batch
            scores = cos_sim(news_embedding, tweet_embeddings)[0]

            # Find indices of matching tweets
            for i, score in enumerate(scores):
                if score > 0.5 and df.at[tweet_indices[i], 'mark_obj'] == 0:
                    df.at[tweet_indices[i], 'mark_obj'] = 1

            print(f"  > matched {sum(scores > 0.5)} tweets")

    return df

if __name__ == "__main__":
    project_path = '/home/xiehh/workspace/Sentence-BERT/tweet_analysis/'

    # Please replace it with the actual Excel file path.
    excel_file_path = project_path + 'HIC/ALL_Tweets_list_HIC_KeyInfo.xlsx'
    # Please replace it with the actual news folder path.
    news_folder_path = project_path + 'HIC/0.HAMAS_News'

    df = read_excel_file(excel_file_path)
    news_dict = read_news_files(news_folder_path)
    df = mark_tweets(df, news_dict)

    # Save the labeled DataFrame to the original Excel file.
    df.to_excel(excel_file_path, index=False)
