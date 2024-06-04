# main.py

import newparser
import pdfextractor
import groqparser
import calculating_balances

def run_all_parsers():
    newparser.main()
    pdfextractor.main()
    groqparser.main()
    calculating_balances.main()

if __name__ == "__main__":
    run_all_parsers()
