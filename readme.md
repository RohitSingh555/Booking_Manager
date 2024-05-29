## How to run these scripts

1. Create virtual environment with this command: `python -m venv env`
2. Activate the environment  `.\env\Scripts\activate.ps1`
3. Install requirements.txt  `pip install -r requirements.txt`
4. Modify the model name according the models that you have currently it will be set to **llama3**
5. Please Move your pdfs into the client_docs folder to let the script extract text from them.
6. And Just run `python run_all.py` in the terminal
7. After running this you will have to input the number of accounts(1,etc), account types(credit or Bank or debit) then account balance(can sum up two to three accounts into one)
8. Make sure to run the chroma db before you run the below command, run the chroma db with this command
   `chroma run --host localhost --port 8000 --path ../vectordb-stores/chromadb`
9. After that run `python run_chatbot.py` in the terminal
10. You will have to enter the prompts then to get a desired output
