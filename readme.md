## How to run these scripts

1. Create virtual environment with this command: `python -m venv env`
2. Activate the environment  `.\env\Scripts\activate.ps1`
3. Install requirements.txt  `pip install -r requirements.txt`
4. Modify the llama model name according to the model name that you have on your pc, currently it is set to **llama3** it's in chatbot.py.
5. Please Move your pdfs into the client_docs folder to let the script extract text from them.
6. Make sure to run the chroma db before you run the below command, run the chroma db with this command in a different terminal.
   `chroma run --host localhost --port 8000 --path ../vectordb-stores/chromadb'`
7. And Just run `python run_all.py` in the terminal
8. After this run `calculating_balances.py`
9. After running this you will have to input the number of accounts(1,etc), account types(credit or Bank or debit) then account balance(can sum up two to three accounts into one)
10. In the main terminal cd into ollama directory `cd ollama`
11. You can run

    ```
    python run_chatbot.py
    ```
//This is to ask questions from the AI bot
    OR
12. After that run `python search.py` in the terminal
13. Follow it up by running the `python chatbot.py`
14. You will have to enter the prompts then to get a desired output

***Note: Once a document is marked as '-read' at the end of the filename it won't be stored as a vectorstore in the database.***
