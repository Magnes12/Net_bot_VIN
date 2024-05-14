import os
import openpyxl
import time

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Read from Excel
wb = openpyxl.load_workbook('cars.xlsx')
sheet = wb['Sheet1']

# Define working columns
vin_column = sheet['A']
kat_column = sheet['B']
typ_column = sheet['C']
rodzaj_column = sheet['D']
data_column = sheet['E']

# Web drivers
edge_options = Options()
# # No show options
# edge_options.add_argument("--headless") options=edge_options

driver = webdriver.Edge()
wait = WebDriverWait(driver, 20)

# # If needed log and pass
LOGIN = os.environ.get("LOGIN")
PASSWORD = os.environ.get("PASSWORD")

try:

    driver.get('https://prod.core.public.vedoc.i.mercedes-benz.com/ui/homepage.html')

    login = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/main/div/div[1]/div/div/div[1]/form/div[2]/div[1]/div/input")))
    login.send_keys(str(LOGIN))
    login.send_keys(Keys.ENTER)
    time.sleep(1)

    password = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[1]/main/div/div[1]/div/div/div[1]/form/div[2]/div[1]/div[1]/input[1]")))
    password.send_keys(str(PASSWORD))
    password.send_keys(Keys.ENTER)
    time.sleep(1)

    driver.get('https://prod.core.public.vedoc.i.mercedes-benz.com/ui/VehicleArrangement.html')

except Exception as e:
    print(f"Error: {e}")


kategoria_info = ""
typ_info = ""
rodzaj_info = ""
data_info = ""

vin_input = wait.until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div/div[2]/form/div[2]/div[4]/div/input")))

# Main loop
for i in range(1, len(vin_column)):

    vin = vin_column[i].value
    auto_info = kat_column[i].value
    print(vin)

    if vin and not auto_info:

        try:

            vin_input.clear()
            vin_input.send_keys(vin)
            vin_input.send_keys(Keys.ENTER)
            time.sleep(0.5)

            wait.until(
                EC.invisibility_of_element_located((By.ID, "loading-bar-spinner"))
            )
            wait.until(
                EC.invisibility_of_element_located((By.ID, "loading-bar"))
            )

            kategoria = wait.until(
                EC.presence_of_element_located((By.XPATH, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'Category')]")))
            kategoria_info = kategoria.text if kategoria else ""

            if kategoria_info == "Osobowe (0)":

                typ = wait.until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='read-only ng-binding' and @data-ng-bind-html]")))
                typ_info = typ.text if typ else ""

                rodzaj = wait.until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'BodyType')]")))
                rodzaj_info = rodzaj.text if rodzaj else ""

            # elif kategoria_info.startswith("Dostawcze"):

            #     typ = wait.until(
            #         EC.visibility_of_element_located((By.XPATH, "//span[@class='read-only ng-binding' and @data-ng-bind-html and contains(@data-ng-bind-html, '111MIX/L4X2')]")))
            #     typ_info = typ.text if typ else ""

            #     rodzaj = wait.until(
            #         EC.visibility_of_element_located((By.XPATH, "//span[@class='read-only ng-binding' and @data-ng-hide and contains(text(), 'viewControl')]")))
            #     rodzaj_info = rodzaj.text if rodzaj else ""

            else:

                typ_info = ""

                rodzaj_info = ""

                data_info = ""

            data = wait.until(
                    EC.visibility_of_element_located((By.XPATH, "//span[@class='read-only ng-binding' and contains(@data-ng-bind, 'idate')]")))
            data_info = data.text if data else ""

            print(kategoria_info)
            print(typ_info)
            print(rodzaj_info)
            print(data_info)

        except Exception as e:
            print(f"Error: {e}")

        kat_column[i].value = kategoria_info
        typ_column[i].value = typ_info
        rodzaj_column[i].value = rodzaj_info
        data_column[i].value = data_info

# Save in Excel
wb.save('cars.xlsx')

# Quit after work done
driver.quit()
