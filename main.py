# skrypt obsługujący pdf, konwersja pdf -> docx, wyodrębnianie stron, łączenie pdf'ów
import PyPDF2
from os import path
# TODO metoda konwert -> pdf
# TODO uwzględnić położenie pliku/ścieżkę
# TODO metoda wpisująca numery stron do listy
# TODO dodać instrukcję programu oraz wypisywanie statusu działań

# może zrobię, metoda sprawdzająca czy rozszerzenie jest poprawnę
# def if_allowed(file_name, action):
#     pass


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
nums = input("podaj numery: ")
print(check_numbers(nums))
action_names = {"1": "merge", "2": "copy pages", "3": "convert to docx"}
print("Wybierz akcję: ")
for el in action_names.keys():
    print(f"{el} - {action_names[el]}")
action = input('... ')
# number_list = [0, 2]

# merge
if action == '1':
    print(action_names[action])
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
    print(action_names[action])
    print("")
    numbers = input("Podaj numery stron. [Możesz podać je po przecinku albo w przedziałach np. 1, 3, 4, 6-9]")
    # number_list = check_numbers(numbers)
    number_list = [0, 2]
    copy_pages(r"C:\Users\gajda\OneDrive\Pulpit\1.pdf", 'plik.pdf', number_list)
# convert_from
elif action == '3':
    print(action_names[action])
    print("Podaj nazwę pliku, który chcesz przekonwertować")
    convert_from(r"C:\Users\gajda\OneDrive\Pulpit\ex3.pdf")


