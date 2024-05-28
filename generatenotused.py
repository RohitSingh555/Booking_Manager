import os
import pandas as pd
from PyPDF2 import PdfReader
import openpyxl
import json
from datetime import datetime
from dotenv import load_dotenv
from pinecone import Pinecone, ServerlessSpec
from transformers import AutoModel, AutoTokenizer
import requests

# Load environment variables
load_dotenv()

# Pinecone configuration
pinecone_api_key = os.getenv("PINECONE_API_KEY")
pinecone_environment = os.getenv("PINECONE_ENVIRONMENT")

# Connect to Pinecone
pc = Pinecone(api_key=pinecone_api_key)

# Create Pinecone index if it doesn't exist
index_name = "excel-generator"
if index_name not in pc.list_indexes().names():
    pc.create_index(
        name=index_name,
        dimension=1536,  # Update this to match the dimension of your embeddings
        metric='euclidean',
        spec=ServerlessSpec(
            cloud='aws',
            region='us-west-2'  # Update to your desired region
        )
    )

pinecone_index = pc.index(index_name)

# Categories
categories = [
    "Income", "Expenses", "Business Expenses", "Tax Deductible Expenses",
    "Recurring Expenses", "Uncertain Expenses"
]

# Initialize state management
state_file = 'processed_files_state.json'
if os.path.exists(state_file):
    with open(state_file, 'r') as file:
        processed_files = json.load(file)
else:
    processed_files = []

# Function to extract text from PDF
def extract_text_from_pdf(pdf_file):
    text = ""
    with open(pdf_file, "rb") as file:
        reader = PdfReader(file)
        for page in reader.pages:
            text += page.extract_text()
    return text

from groq.llmcloud import Completion
# Function to classify transactions using LLaMA3 model
def classify_transaction(description):
    # Replace with actual model call
    with Completion() as completion:
        prompt = f"Classify the following transaction: {description}"
        response, id, stats = completion.send_prompt("llama2-70b-4096", user_prompt=prompt)
        if response != "":
            print(f"\nPrompt: {prompt}\n")
            print(f"Request ID: {id}")
            print(f"Output:\n {response}\n")
            print(f"Stats:\n {stats}\n")
            category = response    
            return category

# Function to read files and categorize transactions
def process_files(input_folder):
    transactions = []
    
    for root, _, files in os.walk(input_folder):
        for file in files:
            file_path = os.path.join(root, file)
            if file_path in processed_files:
                continue

            # Process CSV files
            if file.endswith('.csv'):
                df = pd.read_csv(file_path)
                for _, row in df.iterrows():
                    transactions.append((row['Date'], row['Description'], row['Amount']))
            
            # Process TXT files
            elif file.endswith('.txt'):
                with open(file_path, 'r') as f:
                    lines = f.readlines()
                for line in lines:
                    try:
                        date, desc, amount = line.split(',')
                        transactions.append((date, desc.strip(), float(amount.strip())))
                    except ValueError:
                        continue
            
            # Process PDF files
            elif file.endswith('.pdf'):
                text = extract_text_from_pdf(file_path)
                lines = text.split('\n')
                for line in lines:
                    try:
                        date, desc, amount = line.split(',')
                        transactions.append((date, desc.strip(), float(amount.strip())))
                    except ValueError:
                        continue
            
            processed_files.append(file_path)
    
    # Save the state of processed files
    with open(state_file, 'w') as file:
        json.dump(processed_files, file)
    
    return transactions

# Function to create Excel sheet
def create_excel(transactions):
    with pd.ExcelWriter('expenses_summary.xlsx', engine='openpyxl') as writer:
        for category in categories:
            df = pd.DataFrame(columns=["Date", "Description", "Amount"])
            df.to_excel(writer, sheet_name=category, index=False)

        # Separate transactions by category and write to the respective sheet
        categorized_data = {category: [] for category in categories}
        
        for date, desc, amount in transactions:
            category = classify_transaction(desc)
            if category not in categories:
                category = "Uncertain Expenses"
            categorized_data[category].append({"Date": date, "Description": desc, "Amount": amount})
        
        for category, data in categorized_data.items():
            if data:
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name=category, index=False)

# Main function to orchestrate the process
def main():
    transactions = process_files('client_docs')
    create_excel(transactions)
    
    # Store embeddings in Pinecone
    embeddings = []
    for category, data in categorized_data.items():
        for row in data:
            embedding = {"id": f"{category}_{row['Date']}", "data": row["Amount"]}
            embeddings.append(embedding)
    
    try:
        result = pinecone_index.upsert(
            vectors=[(embedding["id"], embedding["data"]) for embedding in embeddings]
        )
        print(f"Embeddings stored in Pinecone: {result}")
    except Exception as e:
        print(f"Error storing embeddings: {str(e)}")
    
    print(f"Expense summary has been saved to expenses_summary.xlsx")

if __name__ == "__main__":
    main()
