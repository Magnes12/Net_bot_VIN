import os
import openpyxl
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def fetch_data(driver, vin):
    try:
        vin_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div/div[2]/form/div[2]/div[4]/div/input")))
        vin_input.clear()
        vin_input.send_keys(vin)
        vin_input.send_keys(Keys.ENTER)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[@data-datatype='typ']")))
        kategoria = driver.find_element(By.XPATH, "//span[@data-datatype='kategoria']")
        typ = driver.find_element(By.XPATH, "//span[@data-datatype='typ']")
        rodzaj = driver.find_element(By.XPATH, "//span[@data-datatype='rodzaj']")
        data = driver.find_element(By.XPATH, "//span[@data-datatype='data']")
        return kategoria.text, typ.text, rodzaj.text, data.text
    except Exception as e:
        print(f"Error fetching data for VIN {vin}: {e}")
        return "", "", "", ""


def main():
    wb = openpyxl.load_workbook('cars.xlsx')
    sheet = wb['Sheet1']

    vin_column = sheet['A']
    kat_column = sheet['B']
    typ_column = sheet['C']
    rodzaj_column = sheet['D']
    data_column = sheet['E']

    edge_options = Options()
    driver = webdriver.Edge(options=edge_options)

    LOGIN = os.environ.get("LOGIN")
    PASSWORD = os.environ.get("PASSWORD")

    try:
        driver.get('https://prod.core.public.vedoc.i.mercedes-benz.com/ui/homepage.html')
        login_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div/div[1]/div/div/div[1]/form/div[2]/div[1]/div/input")))
        login_input.send_keys(str(LOGIN))
        login_input.send_keys(Keys.ENTER)

        password_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div/div[1]/div/div/div[1]/form/div[2]/div[1]/div[1]/input[1]")))
        password_input.send_keys(str(PASSWORD))
        password_input.send_keys(Keys.ENTER)

        driver.get('https://prod.core.public.vedoc.i.mercedes-benz.com/ui/VehicleArrangement.html')

        for i in range(1, len(vin_column)):
            vin = vin_column[i].value
            auto_info = kat_column[i].value
            if vin and not auto_info:
                kategoria_info, typ_info, rodzaj_info, data_info = fetch_data(driver, vin)
                kat_column[i].value = kategoria_info
                typ_column[i].value = typ_info
                rodzaj_column[i].value = rodzaj_info
                data_column[i].value = data_info
    except Exception as e:
        print(f"Error: {e}")
    finally:
        wb.save('cars.xlsx')
        driver.quit()

if __name__ == "__main__":
    main()
