print("Starting main.py")
import ollamaa.search as search
print("Imported search")
import ollamaa.chatbot as chatbot
print("Imported chatbot")

def run_all():
    print("Running search")
    search.main()
    print("Running chatbot")
    chatbot.main()

if __name__ == "__main__":
    run_all()