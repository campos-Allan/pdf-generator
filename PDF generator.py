"""updating spreadsheets and exporting them as pdf
+ updating another spreadsheet, copying the update into another file
and then printing the result
"""
import os
from datetime import datetime
import time
import win32com.client
import pyautogui as gui

dia = datetime.now().strftime('%d')
ano = datetime.now().year
mes = datetime.now().strftime('%m')

arq = {'path/sheet1.xlsx': [1, 2, 3, 4, 5, 6],
       'path/sheet2.xlsx': [1, 2],
       'path/sheet3.xlsx': [1],
       'path/sheet4.xlsx': [1],
       'path/sheet5.xlsx': [1],
       'path/sheet6.xlsx': [1, 2, 3],
       'path/sheet7.xlsx': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
       'path/sheet8.xlsx': [1, 2, 3, 4]}

arq_escuros = f'path/extra_sheet{ano}_{mes}'

escuros = {'path/query extra_sheet.xlsx': [1],
           arq_escuros: []}

for file, _ in arq.items():
    if 'REFINARIAS' in file:
        time.sleep(60)
    os.system(f'start "excel" "{file}"')

time.sleep(700)
os.system("taskkill /t /im excel.exe")
time.sleep(10)

for file, index in arq.items():
    WB_PATH = file
    ws_index_list = index
    nome = file.split('/')[-1][:-5]
    dir_fim = f'path\\{nome}'
    dir_dia = f'_{ano}_{mes}_{dia}'
    PATH_TO_PDF = dir_fim+dir_dia+'.pdf'
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    time.sleep(4)
    try:
        print('Start conversion to PDF')
        wb = excel.Workbooks.Open(WB_PATH)
        wb.WorkSheets(ws_index_list).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    except Exception as e:
        print('failed.')
        print(e)
        os.system("taskkill /f /im excel.exe")
    else:
        print('Succeeded.')
    finally:
        try:
            wb.Close(False)
            excel.Quit()
            del excel
            wb = None
            excel = None
            os.system("taskkill /f /im EXCEL.exe")
        except Exception as e:
            print('failed.')
            print(e)
            os.system("taskkill /f /im excel.exe")

time.sleep(10)
input('press for autogui ')
for file, _ in escuros.items():
    if 'query' in file:
        os.system(f'start "excel" "{file}"')
        time.sleep(80)
        gui.hotkey('ctrl', 'c')
    else:
        WB_PATH = file
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(WB_PATH)
        wb.WorkSheets('sheetname')
        time.sleep(8)
        time.sleep(1)
        gui.click(200, 800)
        gui.scroll(10000)
        gui.click(150, 880)
        time.sleep(1)
        gui.hotkey('ctrl', 'v')
        time.sleep(3)
        gui.hotkey('printscreen')
        time.sleep(4)
        gui.click(1190, 865)
        time.sleep(1)
        gui.mouseDown(button='right')
        time.sleep(1)
        gui.moveTo(50, 405, duration=2)
        time.sleep(1)
        gui.mouseUp(button='right')
        time.sleep(1)
        gui.click(1900, 20)
        time.sleep(5)
        gui.click(1900, 20)
