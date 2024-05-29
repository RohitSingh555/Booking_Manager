## How to run these scripts

1. Create virtual environment with this command: `python -m venv env`
2. Activate the environment  `.\env\Scripts\activate.ps1`
3. Install requirements.txt  `pip install -r requirements.txt`
4. Modify the model name according the models that you have currently it will be set to **llama3**
5. Please Move your pdfs into the client_docs folder to let the script extract text from them.
6. And Just run `python run_all.py` in the terminal
7. After this run `calculating_balances.py`
8. After running this you will have to input the number of accounts(1,etc), account types(credit or Bank or debit) then account balance(can sum up two to three accounts into one)
9. Make sure to run the chroma db before you run the below command, run the chroma db with this command in a different terminal.
   `chroma run --host localhost --port 8000 --path ../vectordb-stores/chromadb'`
10. In the main terminal cd into ollama directory `cd ollama`
11. After that run `python search.py` in the terminal
12. Follow it up by running the `python chatbot.py`
13. You will have to enter the prompts then to get a desired output


***Note: Once a document is marked as '-read' at the end of the filename it won't be stored as a vectorstore in the database.***
