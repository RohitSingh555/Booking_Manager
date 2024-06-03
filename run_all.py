# main.py

import newparser
import pdfextractor
import groqparser

def run_all_parsers():
    newparser.main()
    pdfextractor.main()
    groqparser.main()

if __name__ == "__main__":
    run_all_parsers()
