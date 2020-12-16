# IMPORT BIBLIOTEK
import os
from pathlib import Path
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import shutil
import matplotlib.pyplot as plt


# ZMIENNE GLOBALNE DLA PROJEKTU
NAZWA_KSIAZKI = 'ORIEUX_TALLEYRAND'
PLIK_CYTATY = NAZWA_KSIAZKI + '.txt'

# DEFINICJE FUNKCJI

# połączenie z własnym Dyskiem Google i Folderem ze zdjęciami kolejnych stron książki
def polaczenie_google_drive(folder):
    # ustanowienie polaczenia z dyskiem Google
    gauth = GoogleAuth()
    # Tworzymy lokalny webserver do autentykacji
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)

    # szukamy katalogu z książką na Dysku Google
    ksiazka = drive.ListFile({'q': "title='{}'  and mimeType = 'application/vnd.google-apps.folder'"
                             .format(folder)}).GetList()

    # książką pobieramy na dysk lokalny wszystkie pliki pdf
    for item in ksiazka:
        file_list = drive.ListFile({'q': "'{}' in parents ".format(item['id'])}).GetList()
        for file in file_list:
            download_mimetype = 'application/pdf'
            if file['mimeType'] == download_mimetype:
                file.GetContentFile(os.path.join(folder, file['title']))

# zamieniamy pobrane pliki pdf na png
def zamiana_pdf_na_png(folder):
    sciezka = Path(folder)
    if sciezka.is_dir():
        # tworzymy listę plików w folderze
        lista_plikow = [x for x in sciezka.iterdir() if x.is_file()]
        for item in lista_plikow:
            # sprawdzenie czy nasz plik jest pdf-em
            if item.suffix == '.pdf':
                # zwykle jest jedna strona ale przyjmujemy
                # ostrozne zaloczenie ze w jednym pdfie moze byc kilka stron i zamiana kazdej na 'png'
                pages = convert_from_path(item, 500)
                licznik = 0
                for page in pages:
                    page.save(folder + '/' + licznik + '_' +item.stem + '.png', 'PNG')
                    licznik+=1

# przycinamy obraz do fragmentu z zaznaczonym kolorem
def detekcja_konturow(plik):
    # wczytujemy plik
    img = cv2.imread(plik)
    # tworzymy wersję pliku w skali szarosci, przyda sie pozniej
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # tworzymy wersję pliku w zmieniajac domyslny schmat koloru na HSV
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    # wybieramy kanal odpowiadajacy za nasycenie koloru
    nasycenie = hsv[:, :, 1]

    # ustawiamy próg dla nasycenia według ktorego wylapiemy kolorowa czesc
    ret, thresh = cv2.threshold(nasycenie, 40, 255, cv2.THRESH_BINARY)

    # odnajdujemy kontury koloru
    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    c = max(contours, key=cv2.contourArea)

    # wymiary prostokąta z konturu
    x, y, w, h = cv2.boundingRect(c)

    # przyciecie obrazu w skali szarosci do kolorowego fragmentu
    out_gray = img_gray[y:y + h, x:x + w].copy()
    return out_gray

# odczytujemy tekst z obrazu w skali szrosci
def czytanie_tekstu(obraz):

    # redukcja szumu
    img = cv2.medianBlur(obraz, 7)

    # prog dla odcieni szarosci
    ret, th1 = cv2.threshold(img, 127, 255, cv2.THRESH_BINARY)

    # podanie ścieżki do tesseract
    # https://stackoverflow.com/questions/50951955/pytesseract-tesseractnotfound-error-tesseract-is-not-installed-or-its-not-i
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    # rozpoznawanie tekstu z przygotowanego obrazu, z ustawieniem języka polskiego
    cytat = pytesseract.image_to_string(th1, 'pol')

    # zwrócenie gotowego cytatu
    return cytat

# zapis cytatow do pliku tekstowego
def zapisywanie_cytatow(plik, cytat):
    tryb = 'a' if os.path.exists(plik) else 'w'
    with open(plik, tryb,encoding='utf-8') as f:
        f.write(cytat + '\n')

# spakowanie w jedną funkcję przycinania zdjec, czytania tekstu i zapisywania do pliku
def zamiana_png_na_txt(katalog, plik_z_cytatami):
    sciezka = Path(katalog)
    # pobranie listy plikow z katalogu
    if sciezka.is_dir():
        lista_plikow = [x for x in sciezka.iterdir() if x.is_file()]
        for item in lista_plikow:
            # sprawdzenie czy nasz plik to 'png'
            if item.suffix == '.png':
                image = detekcja_konturow(os.path.join(NAZWA_KSIAZKI,item.name))
                cytat = czytanie_tekstu(image)
                zapisywanie_cytatow(plik_z_cytatami,item.name)
                # sprawdzenie czy nie ma jakiś dziwnych znaków po OCR
                zapisywanie_cytatow(plik_z_cytatami, cytat.encode('utf-8','ignore').decode('utf-8'))

# tworzymy katolog dla naszych plików tymczasowych
def tworzenie_katalogu(katalog):
    if os.path.isdir(katalog):
        shutil.rmtree(katalog)
        os.mkdir(katalog)
    else:
        os.mkdir(katalog)

# laczymy sie z dyskiem Google i wysylamy plik zbiorczy z cytatami
def wyslanie_pliku(plik):

    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()  # Creates local webserver and auto handles authentication.
    drive = GoogleDrive(gauth)

    with open(plik, 'r', encoding='utf-8') as f:
        pelny_tekst = f.read()
        plik_na_dysku = drive.CreateFile({'title': plik})
        plik_na_dysku.SetContentString(pelny_tekst,encoding='utf-8')
        plik_na_dysku.Upload()

# WYWOLANIE PROGRAMU
if __name__ == '__main__':
    # stworzenie katalogu na pobierane pliki pdf z dysku Google
    # jeśli istnieje usuwamy go z zawartoscia i tworzymy pusty
    if os.path.isfile(PLIK_CYTATY):
        os.remove(PLIK_CYTATY)
    tworzenie_katalogu(NAZWA_KSIAZKI)

    # pobranie zawartości katalogu na dysku google
    polaczenie_google_drive(NAZWA_KSIAZKI)

    # zamiana pdf-ow na png
    zamiana_pdf_na_png(NAZWA_KSIAZKI)

    # zamiana png na tekst
    zamiana_png_na_txt(NAZWA_KSIAZKI, PLIK_CYTATY)

    # wysylamy na dysk google
    wyslanie_pliku(PLIK_CYTATY)

    print('Wykonano pomyślnie!!!')
