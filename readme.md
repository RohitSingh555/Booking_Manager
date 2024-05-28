## How to run these scripts

1. Create virtual environment with this command: `python -m venv env`
2. Activate the environment  `.\env\Scripts\activate.ps1`
3. Install requirements.txt  `pip install -r requirements.txt`
4. And Just run `python run_all.py` in the terminal
5. Make sure to run the chroma db before you run the below command, run the chroma db with this command
   `chroma run --host localhost --port 8000 --path ../vectordb-stores/chromadb`
6. After that run `python run_chatbot.py` in the terminal
