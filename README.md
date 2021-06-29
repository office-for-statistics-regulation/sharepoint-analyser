## sharepoint-analyser

The aim of this repository is to write code that can parse SharePoint files for relevant information.

### Enabled SharePoint searching

Unforunately, we are unable to use an API or a web scraper to get information directly from SharePoint due to security. Instead we can sync our SharePoint folders to a local computer using the OneDrive application. See [this guide for more information](https://support.microsoft.com/en-us/office/sync-sharepoint-files-and-folders-87a96948-4dd7-43e4-aca1-53f3e18bea9b).

The limitations of this are that this process cannot be easily deployed on cloud infrastructure, and must be run on a local, secure machine.

Be careful not to delete or change files, as this will delete and change files on SharePoint too.

### Keyword search

Our first requirement was to create a keyword search tool so we can analyse whether certain files on SharePoint contain certain keywords.

For example, we might want to see what files contain the word `Transparency`. We might also want to if this is in the title of a file, and/or how many times the file contains the term.

#### Documents

Our first requirement is to search through documents, largely now Microsoft Word `.docx` files. To do this we use the python package called `docx`. To install:

`pip install docx`

The function `keyword_search` in `main.py` then uses the configuration files in `config.py` to search for the term in all documents in a specified folder. For example, for transparency, for two dummy documents we get:

| document_name | document_name_contains_word | count_of_word_in_document |
| -- | -- | -- |
| This one doesnt.docx | False | 0 |
| Transparency document.docx | True | 3 |