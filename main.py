import glob, os
from docx import Document
import pandas as pd

import config


def keyword_search():
    document_name = []
    document_name_contains_word = []
    count_of_word_in_document = []

    for file in glob.glob("*.docx*"):
        print(f'Opening {file}')
        
        document_name.append(file)
        
        if config.word_to_search in file.lower():
            document_name_contains_word.append(True)
        else:
            document_name_contains_word.append(False)
        
        document = Document(file)
        count_of_word_in_document_tmp = 0
        for paragraph in document.paragraphs:
            words = paragraph.text.split()
            for word in words:
                if config.word_to_search in word.lower():
                    count_of_word_in_document_tmp += 1

        print(f"Found {config.word_to_search} in {file} {str(count_of_word_in_document_tmp)} times")
        count_of_word_in_document.append(count_of_word_in_document_tmp)

    df = pd.DataFrame(list(zip(document_name, document_name_contains_word, count_of_word_in_document)),
                columns =['document_name', 'document_name_contains_word', 'count_of_word_in_document'])
    print(df)


if __name__ == '__main__':
    os.chdir('C:\\Users\\' + config.username + '\\Office for National Statistics\\' + config.folder_to_search)
    keyword_search()