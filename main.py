import openpyxl
from selenium import webdriver
# from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Read from Excel
wb = openpyxl.load_workbook('cars.xlsx')
sheet = wb['Sheet1']

vin_column = sheet['A']
auto_column = sheet['B']


# chrome_options = Options()
# chrome_options.add_argument("--headless") options=chrome_options

driver = webdriver.Chrome()

for i in range(1, len(vin_column)):
    vin = vin_column[i].value
    print(vin)
    if vin:
        try:
            driver.get(f'path')
            info = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div[3]/div[3]/div[5]/div[1]/p[2]"))
            )
            auto_info = info.text if info else ""
        except Exception as e:
            print(f"Error: {e}")
            auto_info = "Not found"

        auto_column[i].value = auto_info

wb.save('cars.xlsx')

driver.quit()
