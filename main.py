import os
from datetime import datetime
import re

from docx import Document
import pandas as pd
import extract_msg
from tqdm import tqdm

import config


def check_dates_match(document_name, date_created, date_modified):
    if len(document_name) > 1 and document_name[-1] == document_name[-2]:
        if date_created[-1] != date_created[-2] or date_modified[-1] != date_modified[-2]:
            print("Something has gone wrong, the dates for the same document are not the same.")
            sys.exit()


def getMetaData(doc):
    metadata = {}
    prop = doc.core_properties
    metadata["author"] = prop.author
    metadata["created"] = prop.created
    metadata["identifier"] = prop.identifier
    metadata["modified"] = prop.modified
    return metadata


def document_title_contains_kw(word_to_search, file):
    if word_to_search in file.lower():
        return True
    else:
        return False


def add_to_word_count(words, word_to_search):
    for word in words:
        if word_to_search in word.lower():
            return 1


def find_number_of_files():
    file_count = 0
    for file_type in config.file_types:
        configfiles = [os.path.join(dirpath, f)
                       for dirpath, dirnames, files in os.walk(config.path)
                       for f in files if f.endswith(file_type)]

        for file in configfiles:
            file_count += 1
    return file_count


def keyword_search(file_count):
    document_name = []
    document_name_contains_word = []
    count_of_word_in_document = []
    type_of_file = []
    word_searched = []
    folder_searched = []
    date_created = []
    date_modified = []
    author = []
    errors = []
    pbar = tqdm(total=file_count)

    for file_type in config.file_types:
        configfiles = [os.path.join(dirpath, f)
                       for dirpath, dirnames, files in os.walk(config.path)
                       for f in files if f.endswith(file_type)]

        for file in configfiles:
            pbar.update(1)
            try:
                if file_type == '.docx':
                    document = Document(file)
                    metadata_dict = getMetaData(document)

                    for word_to_search in config.words_to_search:
                        type_of_file.append(file_type)
                        folder_searched.append(config.path)
                        date_created.append(metadata_dict.get('created').strftime('%d/%m/%Y'))
                        if metadata_dict.get('modified') != None:
                            date_modified.append(metadata_dict.get('modified').strftime('%d/%m/%Y'))
                        else:
                            date_modified.append('')
                        author.append(metadata_dict.get('author'))
                        document_name.append(file.replace(config.path + '\\', ''))
                        word_searched.append(word_to_search)

                        check_dates_match(document_name, date_created, date_modified)

                        count_of_word_in_document_tmp = 0

                        document_name_contains_word.append(document_title_contains_kw(word_to_search, file))

                        for paragraph in document.paragraphs:
                            words = paragraph.text.split()
                            for word in words:
                                if word_to_search == word.lower():
                                    count_of_word_in_document_tmp += 1

                        count_of_word_in_document.append(count_of_word_in_document_tmp)

                elif file_type == '.msg':
                    msg = extract_msg.Message(file)
                    sender = msg.sender
                    msg_date = msg.date.replace(',', '')
                    msg_message = msg.body
                    msg_subj = msg.subject

                    msg_text = msg_message + ' ' + msg_subj

                    msg_text = re.sub(r'[^\w\s]', '', msg_text)
                    msg_text = msg_text.replace('\r', ' ')
                    msg_text = msg_text.replace('\n', ' ')
                    words = msg_text.split(' ')
                    words = [word.lower() for word in words]

                    for word_to_search in config.words_to_search:

                        type_of_file.append(file_type)
                        folder_searched.append(config.path)
                        date_created.append(
                            datetime.strptime(msg_date, '%a %d %b %Y %H:%M:%S %z').strftime('%d/%m/%Y'))
                        date_modified.append('')
                        author.append(sender)
                        document_name.append(file.replace(config.path + '\\', ''))
                        word_searched.append(word_to_search)

                        check_dates_match(document_name, date_created, date_modified)

                        count_of_word_in_document_tmp = 0

                        document_name_contains_word.append(document_title_contains_kw(word_to_search, file))

                        for word in words:
                            if word_to_search == word.lower():
                                count_of_word_in_document_tmp += 1

                        count_of_word_in_document.append(count_of_word_in_document_tmp)
            except:
                errors.append(f"Error in opening {file.replace(config.path, '')}")
                continue

    df = pd.DataFrame(list(zip(folder_searched, document_name, author, date_created, date_modified,
                               type_of_file, word_searched, document_name_contains_word, count_of_word_in_document)),
                      columns=['folder', 'document_name', 'author', 'date_created', 'date_modified', 'type_of_file',
                               'word_searched', 'document_name_contains_word', 'count_of_word_in_document'])
    print(df)


    df.to_csv(os.path.join(config.path, 'keywords.csv'))

    errors_df = pd.DataFrame({'error': errors})
    print(errors_df)
    errors_df.to_csv(os.path.join(config.path, 'errors.csv'))


if __name__ == '__main__':
    file_count = find_number_of_files()
    keyword_search(file_count)
