#!/usr/bin/env python

"""gdocs-flash-card-maker.py: Accesses docs file in Google Drive and converts table into a comparable excel file for
  upload to Quizlet.TM or another flash card provider"""


import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd

"""
    TODO:
    Handle tables which have blank cells
    Handle multiple documents
    Programmatic upload of excel file to flash card services for services who allow (Quizlet specifically prohibits)
    Reverse flow - from flash card to table(s) in docs
"""

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/documents.readonly']

# The ID of your document (44 char string in doc url).
DOCUMENT_ID = '1SqhlKxwm5u0MorUa-EU4GmSa8123njPSkgvXDAdKZTo'
doc_id = DOCUMENT_ID
scopes = SCOPES
# Output file.
PATH_TO_OUTPUT = '/Users/ryanlenea/Desktop/foo.xlsx'


def get_doc(scopes: list, doc_id: str) -> dict:
    """login and get doc info."""

    # TODO: accept argument option to not store credentials

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)

    # Retrieve the documents contents from the Docs service.
    document = service.documents().get(documentId=doc_id).execute()
    return document


def get_table_cells(document: dict) -> list:
    """get all values inside doc table(s)."""

    #TODO: make this less awful with list comp, multiple funcs, recursion etc.
    # see tables text extacton here: https://developers.google.com/docs/api/samples/extract-text
    # and file structure here: https://developers.google.com/docs/api/concepts/structure

    doc_content = document.get('body').get('content')

    all_tables = []
    for value in doc_content:
        if 'table' in value:
            table = value.get('table')
            print(table)

            full_table = []
            for row in table.get('tableRows'):
                cells = row.get('tableCells')

                row_cells =[]
                for cell in cells:
                    cell_cont = cell.get('content')

                    full_cell_text = ''
                    for cont in cell_cont:
                        elements = cont.get('paragraph')['elements']

                        for inner in elements:
                            text = inner.get('textRun').get('content')
                            full_cell_text += text

                    row_cells += [full_cell_text]
                full_table += row_cells
            all_tables += full_table

    return  all_tables
#
# def get_tables_text(document: dict):
#     doc_content = document.get('body').get('content')
#
#     tables = get_tables(doc_content)
#
#     rows_content = []
#     for table in tables:
#         rows_content.append(get_rows_content(table))
#
#     cells_content = []
#     for row in rows_content:
#         cells_content.append(get_cells_content(row_content))
#
#     #LEFTOFFHERE this isn't working because with each funciton, it creates anotehr two elements.
#     # tables, rows_content both have two elements, should be able to merge all the rows together, not sure why not working
#
#
#
#
# doc_content = document.get('body').get('content')
#
# def get_tables(doc_content):
#     tables = []
#     for value in doc_content:
#         if 'table' in value:
#             tables.append(value.get('table'))
#     return tables
#
# tables = get_tables(doc_content)
# table = tables[0]
#
# def get_rows_content(tables):
#     rows_content = []
#     for table in tables:
#         content = []
#         for row in table.get('tableRows'):
#             cells = row.get('tableCells')
#             content.append(cells)
#         row_content.append(content)
#     return row_content
#
# rows_content = get_rows_content(table)
# row_content = rows_content[0]
#
# def get_cells_content(row_content):
#     cells_content = []
#     for cell in row_content:
#         cells_content.append(cell.get('content'))
#     return cells_content
#
# cells = get_cells_content(row_content)
# cell_content = cells[0]
#
# def get_cell_text(cell_cont):
#     full_cell_text = ''
#     for cont in cell_cont:
#         elements = cont.get('paragraph')['elements']
#
#         for inner in elements:
#             text = inner.get('textRun').get('content')
#             full_cell_text += text
#     return full_cell_text
#
# cell_text = get_cell_text(cell_content)





https://developers.google.com/docs/api/samples/extract-text
https://developers.google.com/docs/api/samples/extract-text

# table
#     tableRows
#         tableCells
#             content
#                 paragraph
#                     elemetents
#                         textrun
#                             content
#                 paragraph
#                     elements
#                         textrun
#                             content




def make_term_list(table_elements: list) -> list:
    """convert list-of-strings to list-of-two-string-lists (term and definition)."""

    #TODO: consolidate to one loop

    # Lose new line characters.
    text_stripped = []
    for n in table_elements:
        s = n.strip()
        text_stripped.append(s.replace("\n", " "))

    lst_of_lsts = [text_stripped[i:i + 2] for i in range(0, len(text_stripped), 2)]
    return lst_of_lsts

def to_excl(card_lst: list, out_path: str) -> None:
    df = pd.DataFrame.from_records(card_lst)
    df.to_excel(out_path, sheet_name='Sheet1', index=False, header=False)

def main() -> None:
    document = get_doc(SCOPES, DOCUMENT_ID)
    cells = get_table_cells(document)
    terms = make_term_list(cells)
    to_excl(terms, PATH_TO_OUTPUT)

if __name__ == "__main__":
    main()
