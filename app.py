import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_products = workbook["Produtos"]

for row in sheet_products.iter_rows(min_row=2):
    name_product = row[0].value
    pyperclip.copy(str(name_product))
    pyautogui.click(908, 237, duration=1)
    pyautogui.hotkey("ctrl", "v")
    description = row[1].value
    pyperclip.copy(str(description))
    pyautogui.click(901, 307, duration=1)
    pyautogui.hotkey("ctrl", "v")
    category = row[2].value
    pyperclip.copy(str(category))
    pyautogui.click(896, 413, duration=1)
    pyautogui.hotkey("ctrl", "v")
    product_code = row[3].value
    pyperclip.copy(str(product_code))
    pyautogui.click(893, 488, duration=1)
    pyautogui.hotkey("ctrl", "v")
    weight = row[4].value
    pyperclip.copy(str(weight))
    pyautogui.click(884, 556, duration=1)
    pyautogui.hotkey("ctrl", "v")
    dimension = row[5].value
    pyperclip.copy(str(dimension))
    pyautogui.click(890, 631, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.click(899, 676, duration=1)
    sleep(5)
    price = row[6].value
    pyperclip.copy(str(price))
    pyautogui.click(904, 258, duration=1)
    pyautogui.hotkey("ctrl", "v")
    quantity_in_stock = row[7].value
    pyperclip.copy(str(quantity_in_stock))
    pyautogui.click(901, 330, duration=1)
    pyautogui.hotkey("ctrl", "v")
    expiration_date = row[8].value
    pyperclip.copy(str(expiration_date))
    pyautogui.click(901, 398, duration=1)
    pyautogui.hotkey("ctrl", "v")
    color = row[9].value
    pyperclip.copy(str(color))
    pyautogui.click(898, 467, duration=1)
    pyautogui.hotkey("ctrl", "v")
    size = row[10].value
    pyautogui.click(921, 538, duration=1)
    if size == "Pequeno":
        pyautogui.click(927, 571, duration=1)
    elif size == "Médio":
        pyautogui.click(919, 594, duration=1)
    else:
        pyautogui.click(924, 618, duration=1)
    meterial = row[11].value
    pyperclip.copy(str(meterial))
    pyautogui.click(912, 607, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.click(915, 655, duration=1)
    sleep(5)
    manufacturer = row[12].value
    pyperclip.copy(str(manufacturer))
    pyautogui.click(913, 275, duration=1)
    pyautogui.hotkey("ctrl", "v")
    country_of_origin = row[13].value
    pyperclip.copy(str(country_of_origin))
    pyautogui.click(913, 345, duration=1)
    pyautogui.hotkey("ctrl", "v")
    observation = row[14].value
    pyperclip.copy(str(observation))
    pyautogui.click(906, 413, duration=1)
    pyautogui.hotkey("ctrl", "v")
    barcode = row[15].value
    pyperclip.copy(str(barcode))
    pyautogui.click(941, 521, duration=1)
    pyautogui.hotkey("ctrl", "v")
    warehouse_location = row[16].value
    pyperclip.copy(str(warehouse_location))
    pyautogui.click(901, 590, duration=1)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.click(907, 639, duration=1)
    pyautogui.click(1274, 177, duration=1)
    pyautogui.click(1086, 457, duration=1)
