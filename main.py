# ############################### #
# by Radosław Tecmer (remceTkedaR)
# radoslaw69tecmer@gmail.com
# ################################ #


# "In the directory na_xx, there are invoices (files) in xls format.
# The program retrieves the files from the specified directory, extracts
# the sum value, and adds up all the sums from the xls files.
# Then it saves the sum value in a txt file in the chosen directory.
# I needed such a program so that I wouldn't have to manually calculate
# the sum of these invoices every time before sending them.
# Because it was important to me that the total sum of these invoices
# does not exceed a specified amount per month."


import os
import xlrd
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askdirectory

# Wybór katalogu z fakturami za pomocą interfejsu graficznego
# Selection of the invoice catalog using the graphical interface
Tk().withdraw()
faktury_dir = askdirectory(title='Wybierz katalog z fakturami INEE')

suma = 0

for root, dirs, files in os.walk(faktury_dir):
    for plik in files:
        if plik.endswith(".xls"):
            plik_path = os.path.join(root, plik)

            try:
                # Otwarcie pliku .xls
                # opening fole .xls
                workbook = xlrd.open_workbook(plik_path)

                # Znalezienie arkusza o nazwie 'Rachunek uproszczony'
                #Finding a sheet named 'Simplified account'
                rachunek_arkusz = None
                for arkusz in workbook.sheet_names():
                    if arkusz == 'Rachunek uproszczony':
                        rachunek_arkusz = workbook.sheet_by_name(arkusz)
                        break

                if rachunek_arkusz:
                    # Pobranie wartości z komórki L20 i dodanie jej do sumy
                    # Retrieve the value from cell L20 and add it to the sum
                    value = rachunek_arkusz.cell_value(19, 11)  # Kolumna L, wiersz 20 (indeksowane od 0)
                    suma += value
                else:
                    print(f"Błąd: Nie znaleziono arkusza 'Rachunek uproszczony' w pliku: {plik_path}")

            except xlrd.XLRDError:
                print(f"Błąd: Nie można otworzyć pliku .xls: {plik_path}")

            # Zamknięcie pliku .xls
            # closing file .xls
            workbook.release_resources()
            del workbook

# Tworzenie nazwy pliku tekstowego ze znacznikiem czasu
# Create a text file name with a timestamp
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
nazwa_pliku = os.path.join(faktury_dir, f"suma_{timestamp}.txt")

# Zapisanie sumy do pliku tekstowego
# Saving the total to a text file
with open(nazwa_pliku, "w") as plik:
    plik.write(str(suma))

# printing
#print("Suma:", suma)

#print("limit: ", '2700')







