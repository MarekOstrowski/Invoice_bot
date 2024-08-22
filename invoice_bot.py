import pyautogui as gui
import time
import subprocess
import pandas as pd
from links import *


def __wait_until_image_is_displayed(image_path, sec):
    # Wait for {sec} seconds at maximum
    for _ in range(sec):
        try:
            point = gui.locateOnScreen(image_path, confidence=0.85)
            if point:
                time.sleep(0.7)
                return point
            
        except gui.ImageNotFoundException:
            pass
        time.sleep(1.0)
    return False


def orders_list():
    orders_ls = pd.read_csv('./orders.txt', sep='\r', header=None)[0].tolist()
    return orders_ls


orders = orders_list()


cords = {"numer_odbiorcy": (654, 240),
         "symbol_link": (531, 297),
         "ZSB_sprzed_wys": (1224, 431),
         "ZSB_dane_wys": (550, 550),
         "FSW_sprzed_wys": (1570, 420),
         "FSW_dane_wys": (850, 510),
         "okno": (775, 76),
         "printer": (121, 113)
         }


def process(order_no):
    gui.hotkey('altleft', '0')  # kartoteka ZSB
    time.sleep(2)
    gui.hotkey('ctrl', 'f12')  # anulowanie filtrowania
    time.sleep(2)
    gui.press('enter')
    time.sleep(3)

    numer_odbiorcy = __wait_until_image_is_displayed(num_odb_png, 16)
    if numer_odbiorcy:
        gui.click(numer_odbiorcy, duration=0.25)
        gui.write(' ' + ' ' + str(order_no), interval=0.25)
        gui.press("enter")
        time.sleep(2)
        gui.moveRel(-130, 60, duration=0.25)
        gui.click()

    zamowienia_sprzedazy = __wait_until_image_is_displayed(zamowienia_sprzedazy_png, 16)
    if zamowienia_sprzedazy:
        print(f"proceeding order no: {order_no}")
        #  time.sleep(13)
        gui.click(cords["ZSB_sprzed_wys"], duration=0.25)
        time.sleep(1)
        gui.click(cords["ZSB_dane_wys"], duration=0.25)
        time.sleep(1)
        gui.press('enter')
        time.sleep(1)
        gui.press('f11')
        time.sleep(1)
        gui.press('enter')

        generuj_dok = __wait_until_image_is_displayed(generuj_dok_png, 16)
        zatwierdzone = __wait_until_image_is_displayed(zatwierdzone_png, 16)
        niezatwierdzdony = __wait_until_image_is_displayed(niezatwierdzony_png, 16)
        potwierdzone = __wait_until_image_is_displayed(potwierdzone_png, 16)

        if generuj_dok and zatwierdzone:
            print(f"Found generuj_dok_png at {generuj_dok}")
            time.sleep(4)
            if gui.click(generuj_dok, duration=0.25):
                print("Clicked generuj_dok")  # seems not to work
            if __wait_until_image_is_displayed(generowanie_dok_png, 16):
                gui.press('enter')
                if __wait_until_image_is_displayed(faktury_sprzedazy_png, 16):  # generowanie faktury sprzedaży

                    fsw_sprzed_wys = __wait_until_image_is_displayed(sprzedaz_wysylkowa_png, 16)
                    if fsw_sprzed_wys:
                        gui.click(fsw_sprzed_wys, duration=0.25)
                    time.sleep(2)
                    fsw_dane_wys = __wait_until_image_is_displayed(dane_wysylkowe_png, 16)
                    if fsw_dane_wys:
                        gui.click(fsw_dane_wys, duration=0.25)
                        time.sleep(3)
                        gui.press('enter')
                        time.sleep(8)  # zatwierdzenie danych wysyłkowych na FSW
                        gui.press('f11')  # zatwierdzenie FSW
                        time.sleep(6)
                        gui.press('enter')
                        time.sleep(6)  # dokument zatwierdzony

        else:
            print(f"NOT Found {generuj_dok_png} at {generuj_dok}")

    #  Drukowanie
    time.sleep(10)  # dokument zatwierdzony
    gui.click(cords["printer"], duration=0.25)
    drukowanie = __wait_until_image_is_displayed(drukowanie_png, 20)
    if drukowanie:
        gui.press(['right', 'up'])
        time.sleep(1)
        gui.hotkey('altleft', 'z')  # zmiana drukarki
        #  time.sleep(6)
        if __wait_until_image_is_displayed(zastosuj_clicked_png, 12):
            #  gui.hotkey('altleft', 'd')  # drukowanie
            drukuj = __wait_until_image_is_displayed(drukuj_png, 12)
            if drukuj:
                gui.click(drukuj)
                zapisywanie_wydruku = __wait_until_image_is_displayed(zapisywanie_wydruku_png, 20)
                #  if zapisywanie_wydruku:
                nazwa_pliku = __wait_until_image_is_displayed(nazwa_pliku_png, 20)
                if nazwa_pliku or zapisywanie_wydruku:
                    gui.write(str(order_no), interval=0.25)   # wpisanie nazwy pliku
                    time.sleep(6)
                    gui.hotkey('altleft', 'z')  # zapisanie .pdf
                    time.sleep(2)
    #  Koniec Drukowania

            for _ in range(3):
                gui.click(cords['okno'], duration=0.25)
                time.sleep(2)
                gui.press(['up', 'enter'])
                time.sleep(2)


def terminal():
    subprocess.call(["cmd", "/c", "start", "/max", terminal_path])
    time.sleep(5)
    # gui.doubleClick(37, 553)
    x, y = gui.locateCenterOnScreen(frog_png, confidence=0.85)
    gui.doubleClick(x, y)
    time.sleep(6)
    gui.press('enter')
    time.sleep(16)


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


terminal_png()


if __wait_until_image_is_displayed(sprzedaz_png, 20):
    print(f"Found {sprzedaz_png} at {gui.locateOnScreen(sprzedaz_png, confidence=0.85)}")

    start_job = time.time()
    for order in orders:
        start = time.time()
        print(f"start of looping item: {order}")
        process(order)
        print(f"end of looping item: {order}")
        end = time.time()
        print(f'Order {order} finished in: ', time.strftime("%H:%M:%S", time.gmtime(end - start)) + '\n')

    print(*orders, sep=", ")
    end_job = time.time()
    print('Job finished in: ', time.strftime("%H:%M:%S", time.gmtime(end_job - start_job)))
    print('\n')
