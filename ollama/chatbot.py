import chromadb
import ollama
import json

# Function to initialize ChromaDB connection
def initialize_chromadb():
    try:
        chroma = chromadb.HttpClient(host="localhost", port=8000)
        return chroma
    except ImportError:
        print("ChromaDB module not found.")
        return None

# Function to retrieve data from ChromaDB based on embeddings
def retrieve_data_from_chromadb(embeddings, collection_name):
    chroma = initialize_chromadb()
    if chroma:
        collection = chroma.get_collection(collection_name)
        result = collection.query(query_embeddings=[embeddings], n_results=1)
        if result and result.get("documents"):
            document = json.loads(result["documents"][0][0])
            return document
    return None

# Function to interact with the chatbot
def chatbot(collection_name):
    while True:
        user_input = input("You: ").strip().lower()

        # You can quit like this if you want to quit
        if user_input in ["exit", "quit", "bye"]:
            print("Chatbot: Goodbye!")
            break

        response = ollama.embeddings(model='nomic-embed-text', prompt=user_input)
        embeddings = response["embedding"]

        # Retrieve data from ChromaDB based on generated embeddings
        document = retrieve_data_from_chromadb(embeddings, collection_name)

        # If data found, print it
        if document:
            print("Chatbot: Here is the information I found:")
            modelquery = f"{user_input} - Answer that question using the following text as a resource and make sure that you always respond in the most human way possible. refrain from giving table like information or json structures, always provide sentences paragraphs and summaries using the resource. Form intelligent sentences giving an impression that you're it's Accountant. \n: {document}"
            stream = ollama.generate(model='llama3', prompt=modelquery, stream=True)
            for chunk in stream:
                if chunk["response"]:
                    print(chunk['response'], end='', flush=True)
                else:
                    print("\n Chatbot: That's all I could find, Please be a little more descriptive for accurate results.")

# Main function
def main():
    collection_name = "buildragwithpython"  
    chatbot(collection_name)

if __name__ == "__main__":
    main()
