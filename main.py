import os
import time
import getpass
import openpyxl

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

clear = lambda: os.system('cls')
clear()
# # If needed log and pass
LOGIN = input("Podaj login: ")
PASSWORD = getpass.getpass("Podaj hasło: ")

# Read from Excel
wb = openpyxl.load_workbook('cars.xlsx')
sheet = wb['Sheet1']
# wb = openpyxl.load_workbook('MODELE - VIN - SŁOWNIK DLA BAZY LC WARSZAWA.xlsx')
# sheet = wb['DANE_VIN']

# Define working columns
vin_column = sheet['A']
kat_column = sheet['B']
typ_column = sheet['C']
rodzaj_column = sheet['D']
data_column = sheet['E']

# Web drivers
edge_options = Options()

driver = webdriver.Edge()
wait = WebDriverWait(driver, 20)

# Try to log in
try:

    print("Otwieram stronę logowania...")
    driver.get('https://prod.core.public.vedoc.i.mercedes-benz.com/ui/homepage.html')

    print("Wpisuje login...")
    login = wait.until(
        EC.element_to_be_clickable((
                By.ID, "userid"
                ))
        )
    login.send_keys(str(LOGIN))
    login.send_keys(Keys.ENTER)
    time.sleep(1)

    print("Wpisuje hasło...")
    password = wait.until(
        EC.element_to_be_clickable((
                By.ID, "password"
                ))
        )
    password.send_keys(str(PASSWORD))
    password.send_keys(Keys.ENTER)
    time.sleep(1)

    print("Przechodzę do VeDOC...")
    driver.get('https://prod.core.public.vedoc.i.mercedes-benz.com/ui/VehicleArrangement.html')

except Exception as e:
    clear()
    print(f"Error: {e}")
    print("Błędny login albo hasło")

# Define empty variables
kategoria_info = ""
typ_info = ""
rodzaj_info = ""
data_info = ""

vin_input = wait.until(
                EC.element_to_be_clickable((
                        By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div/div[2]/form/div[2]/div[4]/div/input"
                        ))
                )
clear()
# Main loop len(vin_column)
for i in range(1, len(vin_column)):

    vin = vin_column[i].value
    kat_info = kat_column[i].value

    # Check if there is data
    if vin and not kat_info:

        print(vin)

        try:

            vin_input.clear()
            vin_input.send_keys(vin)
            vin_input.send_keys(Keys.ENTER)
            time.sleep(0.5)

            wait.until(
                EC.invisibility_of_element_located((
                        By.ID, "loading-bar-spinner"
                        ))
            )
            wait.until(
                EC.invisibility_of_element_located((
                        By.ID, "loading-bar"
                        ))
            )

            kategoria = wait.until(
                EC.presence_of_element_located((
                        By.XPATH, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'Category')]"
                        ))
                )
            kategoria_info = kategoria.text if kategoria else ""

            if kategoria_info == "Osobowe (0)":

                typ = wait.until(
                    EC.visibility_of_element_located((
                            By.XPATH, "//span[@class='read-only ng-binding' and @data-ng-bind-html]"
                            ))
                    )
                typ_info = typ.text if typ else ""

                rodzaj = wait.until(
                    EC.visibility_of_element_located((
                            By.XPATH, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'BodyType')]"
                            ))
                    )
                rodzaj_info = rodzaj.text if rodzaj else ""

            elif kategoria_info.startswith("Dostawcze") or kategoria_info.startswith("Ciężarowe"):

                typ = wait.until(
                    EC.visibility_of_element_located((
                            By.XPATH, "//span[contains(@class, 'read-only ng-binding') and contains(@data-ng-bind-html, 'viewDataObject.vehicleModelDesignation.designation.requestedText')]"
                            ))
                    )
                typ_info = typ.text if typ else ""
                rodzaj_info = ""

            elif kategoria_info.startswith("Klasa"):

                typ = wait.until(
                    EC.visibility_of_element_located((
                            By.XPATH, "//span[@class='read-only ng-binding' and contains(@data-ng-bind-html, 'viewDataObject.vehicleModelDesignation.designation.requestedText')]"
                            ))
                    )
                typ_info = typ.text if typ else ""

                rodzaj = wait.until(
                    EC.visibility_of_element_located((
                            By.XPATH, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'viewDataObject.vehicleModelDesignation.bodyType') and not(text()='3')]"
                            ))
                    )
                rodzaj_info = rodzaj.text if rodzaj else ""

            else:

                typ_info = ""
                rodzaj_info = ""

            try:
                data = wait.until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'idate')]"))
                )
                data_info = data.text if data else ""
            except Exception:
                alert = driver.find_elements(By.CSS_SELECTOR, "span.alert-counter.ng-binding.ng-scope.ng-hide")
                if alert:
                    kategoria_info = "Nie znaleziono pojazdu/Brak uprawień"
                    typ_info = ""
                    rodzaj_info = ""
                    data_info = ""
                else:
                    data_info = "Brak daty"

        except Exception:
            kategoria_info = "Nie znaleziono pojazdu/Brak uprawień"
            typ_info = ""
            rodzaj_info = ""
            data_info = ""

        kat_column[i].value = kategoria_info
        typ_column[i].value = typ_info
        rodzaj_column[i].value = rodzaj_info
        data_column[i].value = data_info

        print(kategoria_info)
        print(typ_info)
        print(rodzaj_info)
        print(data_info)
        print("")

    else:
        print(f"{vin} - dane dla tego VIN'u są kompletne")

# Save Excel
wb.save('cars.xlsx')
print("Excel zapisany.")

# Quit after work done
driver.quit()
print("Zamykam program.")
