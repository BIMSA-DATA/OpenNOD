import torch
import torch.nn as nn
import torch.optim as optim
from torch.utils.data import DataLoader, Dataset
from transformers import AutoModel, AutoTokenizer
import numpy as np
import pandas as pd
import os
import random

# Configuration
DEVICE = torch.device("cuda:2" if torch.cuda.is_available() else "cpu")
MODEL_NAME = "sentence-transformers/all-MiniLM-L6-v2"
BATCH_SIZE = 30
EPOCHS = 10
LEARNING_RATE = 1e-5
TEMPERATURE = 1.16 # Contrastive learning temperature

# Load Pre-trained Model and Tokenizer
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
base_model = AutoModel.from_pretrained(MODEL_NAME).to(DEVICE)

# Custom Dataset Class
class SentimentDataset(Dataset):
    def __init__(self, texts, labels):
        self.texts = texts
        self.labels = labels

    def __len__(self):
        return len(self.texts)

    def __getitem__(self, idx):
        return self.texts[idx], self.labels[idx]

class ContrastiveLoss(nn.Module):
    def __init__(self, temperature=TEMPERATURE):
        super(ContrastiveLoss, self).__init__()
        self.temperature = temperature

    def forward(self, embeddings, labels):
        embeddings = nn.functional.normalize(embeddings, p=2, dim=1)

        similarity_matrix = torch.matmul(embeddings, embeddings.T) / self.temperature
        labels_matrix = labels.unsqueeze(0) == labels.unsqueeze(1)
        labels_matrix = labels_matrix.float()

        weights = torch.tensor(
            [[1.5 if abs(y_i - y_j) == 1 else 1 for y_j in labels.cpu().numpy()] for y_i in labels.cpu().numpy()],
            device=DEVICE)

        positive_mask = labels_matrix - torch.eye(labels_matrix.shape[0], device=DEVICE)
        positives = (similarity_matrix * positive_mask * weights).sum(dim=1)
        negatives = (similarity_matrix * (1 - labels_matrix))

        if negatives.dim() < 2:
            negatives_sum = negatives.exp().sum()
        else:
            negatives_sum = negatives.exp().sum(dim=1)

        # Add a small constant to avoid a zero input for the logarithm.
        eps = 1e-6  # Increase epsilon
        denominator = torch.clamp(negatives_sum, min=eps)
        loss = -torch.log((positives.exp() + eps) / denominator)
        #loss = -torch.log((positives.exp() + eps) / (negatives_sum + eps)).mean()

        return loss.mean()

# Sentence Embedding Model
class SentimentEncoder(nn.Module):
    def __init__(self, model_name):
        super(SentimentEncoder, self).__init__()
        self.encoder = AutoModel.from_pretrained(model_name)
        self.tokenizer = AutoTokenizer.from_pretrained(model_name)
    
    def forward(self, input_ids, attention_mask):
        outputs = self.encoder(input_ids=input_ids, attention_mask=attention_mask)
        return outputs.last_hidden_state[:, 0, :]

    def encode(self, texts):
        # This method handles tokenization and model inference
        encoding = self.tokenizer(texts, padding=True, truncation=True, return_tensors="pt", add_special_tokens=True)
        input_ids = encoding['input_ids'].to(DEVICE)
        attention_mask = encoding['attention_mask'].to(DEVICE)
        with torch.no_grad():
            embeddings = self.encoder(input_ids=input_ids, attention_mask=attention_mask).last_hidden_state[:, 0, :]
        return embeddings

def load_training_data(file_path):
    df = pd.read_excel(file_path)
    texts = df.iloc[:, 2].tolist()  # tweet_text column
    labels = df.iloc[:, 4].tolist()  # label column
    
    # Manually define the label mapping
    label_mapping = {
        'na': 0,
        'impossible': 1,
        'very negative': 1,
        'unlikely': 2,
        'mild negative': 2,
        'not sure': 3,
        'neutral': 3,
        'possible': 4,
        'mild positive': 4,
        'certain': 5,
        'very positive': 5
    }

    # Convert labels to numbers based on manual mapping
    labels = [label_mapping[label] for label in labels]

    return texts, labels, label_mapping

# Function for Contrastive Representation Learning
def train_contrastive_model(train_texts, train_labels):
    # Shuffle training data before creating the dataset
    combined = list(zip(train_texts, train_labels))
    random.shuffle(combined)
    train_texts, train_labels = zip(*combined)  # Unzip after shuffling
    
    train_dataset = SentimentDataset(list(train_texts), list(train_labels))
    train_dataloader = DataLoader(train_dataset, batch_size=BATCH_SIZE, shuffle=True)  # Still keep shuffle=True

    encoder = SentimentEncoder(MODEL_NAME).to(DEVICE)
    contrastive_loss = ContrastiveLoss().to(DEVICE)
    optimizer = optim.AdamW(encoder.parameters(), lr=LEARNING_RATE)

    encoder.train()
    for epoch in range(EPOCHS):
        total_loss = 0
        for texts, labels in train_dataloader:
            inputs = tokenizer(texts, padding=True, truncation=True, return_tensors="pt").to(DEVICE)
            labels = torch.tensor(labels, dtype=torch.long, device=DEVICE)
                        
            optimizer.zero_grad()
            embeddings = encoder(input_ids=inputs['input_ids'], attention_mask=inputs['attention_mask'])
            
            loss = contrastive_loss(embeddings, labels)
            loss.backward()
            # Apply Gradient Clipping to Prevent Instability
            torch.nn.utils.clip_grad_norm_(encoder.parameters(), max_norm=0.5)
            optimizer.step()
            
            total_loss += loss.item()
        print(f"Epoch {epoch + 1}, Loss: {total_loss / len(train_dataloader):.4f}")
    
    return encoder

# Function to save the trained encoder to local disk
def save_encoder_to_local(encoder, save_directory="saved_model"):
    # Ensure the save directory exists
    os.makedirs(save_directory, exist_ok=True)
    
    # Save the model and tokenizer
    encoder.encoder.save_pretrained(save_directory)
    tokenizer.save_pretrained(save_directory)
    print(f"Model saved to {save_directory}")

if __name__ == '__main__':
    # Example Usage
    file_path = "datasets//labeled+para-AGI-600.xlsx"
    #file_path = "datasets//labeled+para-HIC-Israel-600.xlsx"
    #file_path = "datasets//labeled+para-HIC-Hamas-600.xlsx"
    train_texts, train_labels, label_mapping = load_training_data(file_path)

    encoder = train_contrastive_model(train_texts, train_labels)
    # Example: Save the trained encoder to local
    save_encoder_to_local(encoder, save_directory="saved_model//AGI")
