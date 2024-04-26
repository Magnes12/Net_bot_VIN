import os
import openpyxl

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Read from Excel
wb = openpyxl.load_workbook('EXCEL OR EXCEL PATH')
sheet = wb['Sheet1']

# Define working columns
vin_column = sheet['A']
auto_column = sheet['B']

# # Web drivers
# edge_options = Options()
# # No show options
# edge_options.add_argument("--headless") options=edge_options

driver = webdriver.Edge()

# # If needed log and pass
# LOGIN = os.environ.get("LOGIN")
# PASSWORD = os.environ.get("PASSWORD")


# Main loop
for i in range(1, len(vin_column)):
    vin = vin_column[i].value
    auto_info = auto_column[i].value
    print(vin)
    if vin and not auto_info:
        try:
            driver.get(f'HTTP_PATH?q={"VIN"}')
            info = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "XPATH")))
            auto_info = info.text if info else ""
        except Exception as e:
            print(f"Error: {e}")
            auto_info = "Nie znaleziono/Błąd"
        auto_column[i].value = auto_info

# Save in Excel
wb.save('EXCEL OR EXCEL PATH')

# Quit after work done
driver.quit()
