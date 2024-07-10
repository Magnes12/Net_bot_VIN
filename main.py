import os
import sys
import time
import getpass
import openpyxl
import subprocess
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from contextlib import contextmanager


@contextmanager
def suppress_output():
    with open(os.devnull, 'w') as devnull:
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = devnull
        sys.stderr = devnull
        try:
            yield
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr


def clear_console():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')


def load_excel(file_name):
    try:
        wb = openpyxl.load_workbook(file_name)
        sheet = wb['DANE_VIN']
        return wb, sheet
    except Exception:
        print(f"""
        Błąd otwiarnia pliku
        Sprawdź czy plik nazywa się - {file_name}
        Jest zamknięty
        Oraz skoroszyt nazywa się 'DANE_VIN'
        """)
        sys.exit()


def setup_webdriver():
    edge_options = Options()
    edge_options.add_argument('--log-level=3')
    edge_options.add_argument('--log-level=3')
    driver = webdriver.Edge(options=edge_options)
    wait = WebDriverWait(driver, 30)
    driver.maximize_window()
    return driver, wait


def login(driver, wait, login_url, username, password):
    try:
        print("Otwieram stronę logowania...")
        driver.get(login_url)

        print("Wpisuje login...")
        login_input = wait.until(EC.element_to_be_clickable(
            (By.ID, "userid")
            ))
        login_input.send_keys(username)
        login_input.send_keys(Keys.ENTER)
        time.sleep(2)

        print("Wpisuje hasło...")
        password_input = wait.until(EC.element_to_be_clickable(
            (By.ID, "password")
            ))
        password_input.send_keys(password)
        password_input.send_keys(Keys.ENTER)
        time.sleep(2)
    except Exception as e:
        print(f"Błąd logowania: {str(e)}")
        driver.quit()
        sys.exit()


def navigate_to_vedoc(driver, wait, vedoc_url):
    try:
        print("Przechodzę do VeDOC...")
        driver.get(vedoc_url)
        time.sleep(2)
        print("Sprawdzanie czy są wiadomości systemowe...")
        try:
            ok_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[@data-ng-click='okAction($event)']")
                ))
            ok_button.click()
            print("Kliknięto przycisk OK dla wiadomości systemowych")
        except Exception:
            print("Wiadomości systemowe nie pojawiły się")
    except Exception as e:
        print(f"Error: {e} \n Błąd nieznany")
        driver.quit()
        sys.exit()


def process_vins(sheet, driver, wait):
    vin_input = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div/div[2]/form/div[2]/div[4]/div/input")
        ))

    for i in range(1, len(sheet['A'])):
        vin = sheet['A'][i].value
        kat_info = sheet['B'][i].value

        if vin and not kat_info:
            print(f"VIN: {vin}")
            try:
                vin_input.clear()
                vin_input.send_keys(vin)
                vin_input.send_keys(Keys.ENTER)
                time.sleep(0.5)

                wait.until(EC.invisibility_of_element_located(
                    (By.ID, "loading-bar-spinner")))
                wait.until(EC.invisibility_of_element_located(
                    (By.ID, "loading-bar")))

                kategoria_info = extract_data(wait, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'Category')]")
                if kategoria_info:
                    typ_info, rodzaj_info, data_info, vin_info, fin_info = extract_vehicle_data(wait, kategoria_info)
                    sheet['B'][i].value = kategoria_info
                    sheet['C'][i].value = typ_info
                    sheet['D'][i].value = rodzaj_info
                    sheet['E'][i].value = data_info
                    sheet['F'][i].value = vin_info
                    sheet['G'][i].value = fin_info
                else:
                    print("Nie znaleziono pojazdu/Brak uprawień")
                    sheet['B'][i].value = "Nie znaleziono pojazdu/Brak uprawień"
            except Exception as e:
                print(f"Błąd {e}")


        else:
            print(f"{vin} - dane dla tego VIN'u są kompletne")


def extract_data(wait, xpath):
    try:
        element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        return element.text if element else ""
    except Exception:
        return ""


def extract_vehicle_data(wait, kategoria_info):
    typ_info = ""
    rodzaj_info = ""
    data_info = ""
    vin_info = ""
    fin_info = ""

    try:
        vin_info = extract_data(wait, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'viewDataObject.vehicle.activeState.vin')]")
        fin_info = extract_data(wait, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'viewDataObject.vehicle.fin')]")

        if kategoria_info == "Osobowe (0)":
            typ_info = extract_data(wait, "//span[@class='read-only ng-binding' and @data-ng-bind-html]")
            rodzaj_info = extract_data(wait, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'BodyType')]")

        elif kategoria_info.startswith("Dostawcze") or kategoria_info.startswith("Ciężarowe"):
            typ_info = extract_data(wait, "//span[contains(@class, 'read-only ng-binding') and contains(@data-ng-bind-html, 'viewDataObject.vehicleModelDesignation.designation.requestedText')]")

        elif kategoria_info.startswith("Klasa"):
            typ_info = extract_data(wait, "//span[@class='read-only ng-binding' and contains(@data-ng-bind-html, 'viewDataObject.vehicleModelDesignation.designation.requestedText')]")
            rodzaj_info = extract_data(wait, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'viewDataObject.vehicleModelDesignation.bodyType') and not(text()='3')]")

        data_info = extract_data(wait, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'idate')]")

    except Exception:
        print("Błąd pobierania danych")

    print(f"{vin_info}\n{fin_info}\n{kategoria_info}\n{typ_info}\n{rodzaj_info}\n{data_info}\n")
    return typ_info, rodzaj_info, data_info, vin_info, fin_info


def main():
    clear_console()

    excel_file = "NUMERY_VIN.xlsx"
    wb, sheet = load_excel(excel_file)

    with suppress_output():
        driver, wait = setup_webdriver()

    clear_console()
    login_url = 'https://prod.core.public.vedoc.i.mercedes-benz.com/ui/homepage.html'
    vedoc_url = 'https://prod.core.public.vedoc.i.mercedes-benz.com/ui/VehicleArrangement.html'

    clear_console()
    username = input("Podaj login: ")
    password = getpass.getpass("Podaj hasło: ")

    clear_console()

    login(driver, wait, login_url, username, password)
    navigate_to_vedoc(driver, wait, vedoc_url)
    process_vins(sheet, driver, wait)

    wb.save(excel_file)
    print("Excel zapisany.")
    driver.quit()

    print("Otwieram Excel...")
    if sys.platform == "win32":
        os.startfile(excel_file)
    else:
        subprocess.call(('open', excel_file))

    for i in range(3, 1, -1):
        print(f"Zamykanie programu...{i}")
        time.sleep(1)
    sys.exit()


if __name__ == "__main__":
    main()
