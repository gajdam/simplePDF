# skrypt obsługujący pdf, konwersja pdf -> docx, wyodrębnianie stron, łączenie pdf'ów
import os.path
import PyPDF2
import PySimpleGUI as sg


def check_numbers(numbers_string):
    parts = numbers_string.split(',')
    result = []
    for part in parts:
        if '-' in part:
            start, end = part.split('-')
            for i in range(int(start), int(end) + 1):
                result.append(i - 1)
        else:
            result.append(int(part) - 1)
    return result


# action 3
# def convert_from(pdf_n):
#     try:
#         with open(pdf_n, 'rb') as f:
#             reader = PyPDF2.PdfReader(f)
#             writer = PyPDF2.PdfWriter()
#             for page in reader.pages:
#                 page_obj = page
#                 writer.add_page(page_obj)
#         output_f = create_file_name(pdf_n) + '.docx'
#         with open(output_f, 'wb') as file_docx:
#             writer.write(file_docx)
#         print("Utworzono plik")
#         sg.popup("File created successfully")
#
#     except Exception as e:
#         print(f'{e}')


def extract_text_from_pdf(pdf_file):
    try:
        with open(pdf_file, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            txt_file = create_file_name(pdf_file) + '.txt'
            with open(txt_file, 'w') as t:
                for page in reader.pages:
                    text = page.extract_text()
                    t.write(text)
        sg.popup("File created successfully")

    except Exception as e:
        print(f'{e}')


# action 2
def extract_pages(input_n, page_numbers):
    try:
        with open(input_n, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            writer = PyPDF2.PdfWriter()
            pages = reader.pages
            for i in range(len(pages)):
                if i in page_numbers:
                    writer.add_page(pages[i])
            output_n = create_file_name(input_n) + '.pdf'
            # print(output_n)
            with open(output_n, 'wb') as file_output:
                writer.write(file_output)

        sg.popup("File created successfully")

    except Exception as e:
        print(f'{e}')


# action 1
def mergePDF(pdf_n):
    try:
        pdf_files = [open(filepdf, 'rb') for filepdf in pdf_n]
        merger = PyPDF2.PdfMerger()
        for file in pdf_files:
            merger.append(file)
        output_f = create_file_name(pdf_n[0]) + '.pdf'
        # print(output_f)
        with open(output_f, 'wb') as f:
            merger.write(f)
        sg.popup("File created successfully")

    except Exception as e:
        print(f'{e}')


def create_file_name(file_path):
    file_name, file_extension = os.path.splitext(os.path.basename(file_path))
    return file_name + '_created'


def get_file_name(file_path):
    return os.path.basename(file_path)


def update_pdf_names(file_paths, input_window):
    pdfs_n = [get_file_name(pdf) for pdf in file_paths]
    element = input_window['pdfs']
    element.Update('; '.join(pdfs_n))


# TODO zabezpieczyć rozszerzenia plików
# TODO zabezpieczyć zakres stron

# layout for main window
layout_main = [
    # [sg.Text('What you want to do?')],
    # [sg.Text('')],
    [sg.Button('Merge PDFs'), sg.Button('Extract pages'), sg.Button('Extract text')],
    [sg.Button('Exit'), sg.Button('About')]
]
window = sg.Window('PDF Manager', layout_main)

while True:
    event, values = window.read()

    if event in (None, 'Exit'):
        break
    # merge
    if event == 'Merge PDFs':
        merge_layout = [[sg.Text('Choose PDF files:')],
                        [sg.Input(key='pdfs', enable_events=True), sg.FilesBrowse()],
                        [sg.Button('Submit')]
                        ]
        merge_window = sg.Window('Merge files', merge_layout)
        pdf_paths = []
        while True:
            merge_event, merge_values = merge_window.read()
            if merge_event in (None, 'Exit'):
                merge_window.close()
                break
            elif merge_event == 'Submit':
                # pdfs = merge_values['pdfs'].split(';')
                # print(pdf_paths)
                mergePDF(pdf_paths)
                sg.popup("ok")
            elif merge_event == 'pdfs':
                pdf_paths = merge_values['pdfs'].split(';')
                # print(pdf_paths)
                update_pdf_names(pdf_paths, merge_window)
    # extract
    if event == 'Extract pages':
        extract_layout = [[sg.Text('Choose PDF file:')],
                          [sg.Input(key='pdfs', enable_events=True), sg.FilesBrowse()],
                          [sg.Input(key='numbers', enable_events=True)],
                          [sg.Button('Submit')]
                          ]

        extract_window = sg.Window('Extract pages', extract_layout)

        pdf_paths = []
        while True:
            extract_event, extract_values = extract_window.read()
            if extract_event in (None, 'Exit'):
                extract_window.close()
                break
            elif extract_event == 'Submit':
                # pdfs = merge_values['pdfs'].split(';')
                print(pdf_paths)
                numbers = str(extract_values['numbers'])
                extract_pages(pdf_paths[0], check_numbers(numbers))
            elif extract_event == 'pdfs':
                pdf_paths = extract_values['pdfs'].split(';')
                print(pdf_paths)
                update_pdf_names(pdf_paths, extract_window)

    if event == 'Extract text':
        convert_layout = [[sg.Text('Choose PDF file:')],
                          [sg.Input(key='pdfs', enable_events=True), sg.FilesBrowse()],
                          [sg.Button('Submit')]
                          ]

        convert_window = sg.Window('Extract text', convert_layout)

        pdf_paths = []
        while True:
            convert_event, convert_values = convert_window.read()
            if convert_event in (None, 'Exit'):
                convert_window.close()
                break
            elif convert_event == 'Submit':
                # pdfs = merge_values['pdfs'].split(';')
                print(pdf_paths)
                extract_text_from_pdf(pdf_paths[0])
            elif convert_event == 'pdfs':
                pdf_paths = convert_values['pdfs'].split(';')
                print(pdf_paths)
                update_pdf_names(pdf_paths, convert_window)

    if event == 'About':
        about_layout = [[sg.Text('Python project')],
                        [sg.Text('Simple PDF manager')],
                        [sg.Text('Version: 1.0')],
                        [sg.Text('Author: Maciej Gajda')]
                        ]
        about_window = sg.Window('About', about_layout)
        about_event, about_values = about_window.read()
        if about_event in (None, 'Exit'):
            about_window.close()

