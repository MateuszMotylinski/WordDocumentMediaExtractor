import WordExtractor
import os

# Get the path of the folder where the script is located
script_folder = os.path.dirname(os.path.abspath(__file__))

# Initialize an empty list to store Word documents
word_docs = []

# Loop through files in the script folder
for filename in os.listdir(script_folder):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        word_docs.append(filename)
        
parser = WordExtractor.DocxToJsonParser()
#parser.SetDocxFilesList(word_docs)
parser.ParseAllProjectPagesDocxToJson()

input("Press Enter to exit...")