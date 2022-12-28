# skrypt obsługujący pdf, konwersja pdf -> docx, wyodrębnianie stron, łączenie pdf'ów
import PyPDF2
from os import path
# TODO metoda wyodrębnienie stron
# TODO metoda konwert -> pdf
# TODO uwzględnić położenie pliku/ścieżkę
# TODO metoda wpisująca numery stron do listy
# TODO dodać instrukcję programu oraz wypisywanie statusu działań

# może zrobię, metoda sprawdzająca czy rozszerzenie jest poprawnę
# def if_allowed(file_name, action):
#     pass


def check_numbers(numbers):
    lista = []
    return lista


# TODO nazwa output
# action 3
def convert_from(pdf_n):
    with open(pdf_n, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        writer = PyPDF2.PdfWriter()
        for page in reader.pages:
            page_obj = page
            writer.add_page(page_obj)

    with open("output.docx", 'wb') as file_docx:
        writer.write(file_docx)
    print("Utworzono plik")


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
    print("dzienki dziala")


# TODO uporządkowac nazwy
# TODO wyjątek do sprawdzania utworzenia pliku
# action 1
def mergePDF(pdf_n, output_f):

    pdf_files = [open(filepdf, 'rb') for filepdf in pdf_n]
    merger = PyPDF2.PdfMerger()
    # writer = PyPDF2.PdfWriter()
    for file in pdf_files:
        merger.append(file)

    with open(output_f, 'wb') as f:
        merger.write(f)
    print(f"Utworzono plik: {output_f}")


# TODO nazwa wybranej aktywności (słownik)
# TODO poprawić pętle i dodać zewnętrzną while action !q
# TODO dokończyć action 2
# TODO dokończyć action 3
# TODO zabezpieczyć rozszerzenia plików
# TODO zabezpieczyć zakres stron
# control
print("Wybierz akcję: ")
action = input("1 - merge\n2 - wyodrębnij strony\n")
# number_list = [0, 2]

# merge
if action == '1':
    pdf_names = []
    output_file = ""
    while True:
        file = input("Podaj nazwę pliku lub 'q' jeśli chcesz zakończyć: ")
        file_name = path.basename(file).split('.')[0]
        if file.lower() == 'q':
            if not pdf_names:
                print("Nie podano żadnego pliku!")
                continue
            break
        pdf_names.append(file)
        output_file += "".join(file_name)
    output_file += "".join('.pdf')
    print(output_file)
    mergePDF(pdf_names, output_file)

# copy_pages
elif action == '2':
    # numbers = input("Podaj numery stron")
    # number_list = check_numbers(numbers)
    number_list = [0, 2]
    copy_pages(r"C:\Users\gajda\OneDrive\Pulpit\1.pdf", 'plik.pdf', number_list)
# convert_from
elif action == '3':
    convert_from(r"C:\Users\gajda\OneDrive\Pulpit\ex3.pdf")


