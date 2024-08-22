import pyautogui as gui
import time
import subprocess
import pandas as pd
from links import *


cords = {"numer_odbiorcy": (654, 240),
         "symbol_link_ZSB": (525, 310),
         "ZSB_sprzed_wys": (1224, 431),
         "ZSB_dane_wys": (550, 550),
         "FSW_sprzed_wys": (1570, 420),
         "FSW_dane_wys": (850, 510),
         "okno": (775, 76),
         "printer": (121, 113)
         }


def invoice_list():
    invoices_ls = pd.read_csv('./invoices.txt', sep=',', header=None)[2].tolist()
    return invoices_ls


invoices = invoice_list()


def __wait_until_image_is_displayed(image_path, sec):
    # Wait for {sec} seconds at maximum
    for _ in range(sec):
        try:
            point = gui.locateOnScreen(image_path, confidence=0.85)
            if point:
                time.sleep(0.1)
                return point
        except gui.ImageNotFoundException:
            pass
        time.sleep(1.0)
    return False


def terminal_png():
    subprocess.call(["cmd", "/c", "start", "/max", terminal_path])
    time.sleep(5)

    point = __wait_until_image_is_displayed(frog_png, 8)
    if point:
        print(f"Found {frog_png} at {point}")
        gui.doubleClick(point, duration=0.5)
        print("pressed enter")

        if __wait_until_image_is_displayed(logowanie_png, 6):
            print(f"Found {logowanie_png} at {gui.locateOnScreen(logowanie_png, confidence=0.85)}")
            gui.press('enter')
            print("pressed enter")


def process_cn(invoice_no):
    gui.hotkey('altleft', '1')  # kartoteka FSN
    time.sleep(2)
    gui.hotkey('ctrl', 'f12')  # kasowanie filtra
    time.sleep(2)
    gui.press('enter')
    time.sleep(3)

    symbol = __wait_until_image_is_displayed(symbol_png, 16)
    if symbol:
        print(f"Found {symbol_png} at {symbol}")
        gui.click(symbol, duration=0.25)
        gui.write(' ' + ' ' + str(invoice_no), interval=0.25)
        gui.press("enter")
        time.sleep(2)
        gui.moveRel(-30, 60, duration=0.25)
        gui.click()

    skoryguj = __wait_until_image_is_displayed(skoryguj_png, 16)
    if skoryguj:
        print(f"Found {skoryguj_png} at {skoryguj}")
        gui.click(skoryguj, duration=0.25)

        zeruj_ilosci = __wait_until_image_is_displayed(zeruj_ilosci_png, 16)
        if zeruj_ilosci:
            #  gui.press(['down', 'down', 'enter', 'enter'], interval=3)
            gui.click(zeruj_ilosci, duration=0.25)
            yes = __wait_until_image_is_displayed(yes_png, 16)
            if yes:
                gui.press("enter")

    kfs = __wait_until_image_is_displayed(KFS_png, 16)
    if kfs:
        print(f"Found {KFS_png} at {kfs}")

        pozostale = __wait_until_image_is_displayed(pozostale_png, 16)
        if pozostale:
            print(f"Found {pozostale_png} at {pozostale}")
            gui.click(pozostale)

        przyczyna_korekty = __wait_until_image_is_displayed(przyczyna_korekty_png, 16)
        if przyczyna_korekty:
            print(f"Found {przyczyna_korekty_png} at {przyczyna_korekty}")
            gui.click(przyczyna_korekty)
            gui.press(['1', 'enter', 'enter'], interval=3)

        fakture_odebral = __wait_until_image_is_displayed(fakture_odebral_png, 16)
        if fakture_odebral:
            print(f"Found {fakture_odebral_png} at {fakture_odebral}")
            gui.click(fakture_odebral)
            gui.write('*', interval=0.25)

        sprzedaz_wysylkowa = __wait_until_image_is_displayed(sprzedaz_wysylkowa_png, 16)
        if sprzedaz_wysylkowa:
            print(f"Found {sprzedaz_wysylkowa_png} at {sprzedaz_wysylkowa}")
            gui.click(sprzedaz_wysylkowa)

        pobierz_dane = __wait_until_image_is_displayed(pobierz_dane_png, 16)
        if pobierz_dane:
            print(f"Found {pobierz_dane_png} at {pobierz_dane}")
            gui.click(pobierz_dane)
            time.sleep(10)
            gui.press('f11')
            pytanie_czy_zatwierdzic = __wait_until_image_is_displayed(pytanie_czy_zatwierdzic_png, 16)
            if pytanie_czy_zatwierdzic:
                gui.press('enter')

        #  Drukowanie
        time.sleep(10)  # dokument zatwierdzony
        gui.click(cords["printer"], duration=0.25)
        drukowanie = __wait_until_image_is_displayed(drukowanie_png, 20)
        if drukowanie:
            gui.press(['right', 'up'])
            time.sleep(1)
            gui.hotkey('altleft', 'z')  # zmiana drukarki
            #  time.sleep(8)
            if __wait_until_image_is_displayed(zastosuj_clicked_png, 12):
                #  gui.hotkey('altleft', 'd')  # drukowanie
                drukuj = __wait_until_image_is_displayed(drukuj_png, 12)
                if drukuj:
                    gui.click(drukuj)
                    zapisywanie_wydruku = __wait_until_image_is_displayed(zapisywanie_wydruku_png, 20)
                    #  if zapisywanie_wydruku:
                    nazwa_pliku = __wait_until_image_is_displayed(nazwa_pliku_png, 20)
                    if nazwa_pliku or zapisywanie_wydruku:
                        gui.write(str(invoice_no).replace('/', '_'), interval=0.25)  # wpisanie nazwy pliku
                        time.sleep(6)
                        gui.hotkey('altleft', 'z')  # zapisanie .pdf
                        time.sleep(2)
        #  Koniec Drukowania

                for _ in range(3):
                    gui.click(cords['okno'], duration=0.25)
                    time.sleep(2)
                    gui.press(['up', 'enter'])
                    time.sleep(2)


terminal_png()
time.sleep(12)

start_job = time.time()
for invoice in invoices:
    start = time.time()
    print(f"start of looping item: {invoice}")
    process_cn(invoice)
    print(f"end of looping item: {invoice}")
    end = time.time()
    print(f'Order {invoice} finished in: ', time.strftime("%H:%M:%S", time.gmtime(end - start)) + '\n')

print(*invoices, sep='\n')
end_job = time.time()
print('Job finished in: ', time.strftime("%H:%M:%S", time.gmtime(end_job - start_job)))
print('\n')
