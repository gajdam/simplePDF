import PyPDF2
from os import path
# TODO metoda wyodrębnienie stron


def mergePDF(pdf_n, output_f):
    #TODO uporządkowac nazwy
    pdf_files = [open(filepdf, 'rb') for filepdf in pdf_n]
    merger = PyPDF2.PdfMerger()
    # writer = PyPDF2.PdfWriter()
    for file in pdf_files:
        merger.append(file)

    with open(output_f, 'wb') as f:
        merger.write(f)
    print(f"Utworzono plik o nazwie: {output_f}")

# def convertPDF()
print("Wybierz akcję: ")
action = input("1 - merge\n2 - wyodrębnij strony\n")
#TODO nazwa wybranej aktywności (słownik)


if action == '1':
    pdf_names = []
    output_file = ""
    while True:
        file = input("Podaj nazwę pliku lub 'q' jeśli chcesz zakończyć: ")
        file_name = path.basename(file).split('.')[0]
        if file.lower() == 'q':
            break
        pdf_names.append(file)
        output_file += "".join(file_name)
    output_file += "".join('.pdf')
    print(output_file)
    mergePDF(pdf_names, output_file)
elif action == '2':
    input_file = input("Wpisz nazwę pliku: [pamiętaj o rozszerzeniu]")
