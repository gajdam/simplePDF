# skrypt obsługujący pdf, konwersja pdf -> docx, wyodrębnianie stron, łączenie pdf'ów
import os.path

import PyPDF2
from os import path
import PySimpleGUI as sg

# TODO metoda konwert -> pdf
# TODO uwzględnić położenie pliku/ścieżkę
# TODO metoda wpisująca numery stron do listy
# TODO dodać instrukcję programu oraz wypisywanie statusu działań


def check_numbers(numbers_string):
    parts = numbers_string.split(',')
    result = []
    for part in parts:
        if '-' in part:
            start, end = part.split('-')
            for i in range(int(start), int(end)+1):
                result.append(i)
        else:
            result.append(int(part))
    return result


# TODO nazwa output
# action 3
def convert_from(pdf_n):
    try:
        with open(pdf_n, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            writer = PyPDF2.PdfWriter()
            for page in reader.pages:
                page_obj = page
                writer.add_page(page_obj)

        with open("output.docx", 'wb') as file_docx:
            writer.write(file_docx)
        print("Utworzono plik")
    except Exception as e:
        print(f'{e}')


# TODO ulepszyć
# TODO nazwa output
# action 2
def copy_pages(input_n, output_n, page_numbers):
    try:
        with open(input_n, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            writer = PyPDF2.PdfWriter()
            pages = reader.pages
            for i in range(len(pages)):
                if i in page_numbers:
                    writer.add_page(pages[i])
            with open(output_n, 'wb') as file_output:
                writer.write(file_output)
    except Exception as e:
        print(f'{e}')
    # print("dzienki dziala")


# TODO uporządkowac nazwy
# TODO wyjątek do sprawdzania utworzenia pliku
# action 1
def mergePDF(pdf_n, output_f):
    try:
        pdf_files = [open(filepdf, 'rb') for filepdf in pdf_n]
        merger = PyPDF2.PdfMerger()
        # writer = PyPDF2.PdfWriter()
        for file in pdf_files:
            merger.append(file)

        with open(output_f, 'wb') as f:
            merger.write(f)
        print(f"Utworzono plik: {output_f}")
    except Exception as e:
        print(f'{e}')


def create_file_name(name_in):
    return path.basename(name_in).split('.')[0].join('_created')


# TODO poprawić pętle i dodać zewnętrzną while action !q
# TODO dokończyć action 2
# TODO dokończyć action 3
# TODO zabezpieczyć rozszerzenia plików
# TODO zabezpieczyć zakres stron

# control
layout_tmp = [
    [sg.Text('Choose PDF files: ')],
    [sg.Input(key='pdfs', enable_events=True), sg.FilesBrowse()],
    [sg.Button('Merge PDFs'), sg.Button('Copy pages')]
]

layout_main = [
    # [sg.Text('What you want to do?')],
    # [sg.Text('')],
    [sg.Button('Merge PDFs'), sg.Button('Extract pages'), sg.Button('Convert PDF')]
]


window = sg.Window('PDF Manager', layout_main)

while True:
    merge_layout = [
        [sg.Text('Choose PDF files:')],
        [sg.Input(key='pdfs', enable_events=True), sg.FilesBrowse()],
        [sg.Button('Merge')]
    ]
    extract_layout = [
        [sg.Text('Choose PDF file and pages:')],
        [sg.Input(key='pdfs', enable_events=True), sg.FilesBrowse()],
        [sg.Button('Extract')]
    ]
    convert_layout = [
        [sg.Text('Convert PDF:')],
        [sg.Input(key='pdfs', enable_events=True), sg.FilesBrowse()],
        [sg.Button('Convert')]
    ]
    event, values = window.read()

    if event in (sg.WIN_CLOSED, 'Anuluj'):
        break

    if event == 'Merge PDFs':
        merge_window = sg.Window('Merge PDFs', merge_layout)
        event, values = merge_window.read()
        if event == 'Merge':
            # mergePDF()
            pass
        elif event == 'pdfs':
            if values['pdfs']:
                pdfs = values['pdfs'].split(';')
                pdfs_names = [os.path.basename(pdf) for pdf in pdfs]
                # window.FindElement('pdfs').Update(';'.join(pdfs_names))
                # window[values['pdfs']].Update(';'.join(pdfs_names))
                input_element = merge_window['pdfs']
                input_element.Update(';'.join(pdfs_names))

    if event == 'Convert PDF':
        convert_window = sg.Window('Convert PDF', convert_layout)
        event, values = convert_window.read()

    if event == 'Extract pages':
        extract_window = sg.Window('Extract Pages', extract_layout)
        event, values = extract_window.read()

