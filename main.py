import openpyxl
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Read from Excel
wb = openpyxl.load_workbook('cars.xlsx')
sheet = wb['Sheet1']

vin_column = sheet['A']
auto_column = sheet['B']


edge_options = Options()
edge_options.add_argument("--headless")

driver = webdriver.Edge(options=edge_options)

for i in range(1, len(vin_column)):
    vin = vin_column[i].value
    print(vin)
    if vin:
        try:
            driver.get(f'https://pl.wikipedia.org/wiki/{vin}')
            info = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div/div[3]/main/div[3]/div[3]/div[1]/h2[1]/span[1]")))
            auto_info = info.text if info else ""
        except Exception as e:
            print(f"Error: {e}")
            auto_info = "Nie znaleziono/Błąd"
        auto_column[i].value = auto_info

wb.save('cars.xlsx')

driver.quit()
