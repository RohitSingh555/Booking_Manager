import chromadb
import ollama
import pandas as pd
import json
import os

# Function to initialize ChromaDB connection
def initialize_chromadb():
    try:
        chroma = chromadb.HttpClient(host="localhost", port=8000)
        print("ChromaDB connection initialized successfully.")
        return chroma
    except Exception as e:
        print(f"Failed to initialize ChromaDB connection: {e}")
        return None

# Function to store Excel data and embeddings in ChromaDB
def store_excel_data_in_chroma(excel_file_path, collection_name, embedmodel='nomic-embed-text'):
    try:
        excel_data = pd.read_excel(excel_file_path, sheet_name=None)
        print("Excel file loaded successfully.")
    except FileNotFoundError:
        print(f"Excel file not found at path: {excel_file_path}")
        return
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    chroma = initialize_chromadb()
    if not chroma:
        return

    try:
        collection = chroma.get_or_create_collection(collection_name)
        print(f"Collection '{collection_name}' accessed or created successfully.")
    except Exception as e:
        print(f"Failed to access or create collection '{collection_name}': {e}")
        return

    for sheet_name, df in excel_data.items():
        try:
            print(f"Processing sheet: {sheet_name}")
            # Convert each DataFrame to a list of dictionaries
            sheet_data = df.to_dict(orient='records')
            sheet_data_json = json.dumps(sheet_data)

            # Use Ollama to generate embeddings for the sheet data
            response = ollama.embeddings(model=embedmodel, prompt=f"Excel data: {sheet_name}")
            embeddings = response["embedding"]

            # Generate unique ID for the document
            doc_id = f'excel_data_{sheet_name}'
            print(f"Generated document ID: {doc_id}")

            # Store serialized data in ChromaDB
            collection.add(
                ids=[doc_id],
                documents=[sheet_data_json],
                embeddings=[embeddings], 
            )
            print(f"Data for sheet '{sheet_name}' stored successfully with embeddings.")
        except KeyError as e:
            print(f"Missing data or invalid format in sheet '{sheet_name}': {e}")
        except Exception as e:
            print(f"Error processing sheet '{sheet_name}': {e}")

# Main function
def main():
    file_path = "ollamaa\categorized_data.xlsx"
    collection_name = "buildragwithpython"

    # # Check if file exists
    # if os.path.exists(file_path):
    print("File exists! Proceeding with data storage.")
    store_excel_data_in_chroma(file_path, collection_name)
    # else:
    #     print("File does not exist at the specified path.")

if __name__ == "__main__":
    main()
