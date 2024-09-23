import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    pyautogui.click(1557,457,duration=1.5)
    pyautogui.write(linha[0].value)

    pyautogui.click(1563,482,duration=1.5)
    pyautogui.write(linha[1].value)

    pyautogui.click(1579,509,duration=1.5)
    pyautogui.write(str(linha[2].value))

    pyautogui.click(1659,534,duration=1.5)
    pyautogui.write(linha[3].value)

    pyautogui.click(1509,562,duration=1.5)
    
    pyautogui.click(819,567,duration=1.5)

