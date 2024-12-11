"""updating spreadsheets and exporting them as pdf
+ updating another spreadsheet, copying the update into another file
and then printing the result
"""
import os
from datetime import datetime
import time
import win32com.client
import pyautogui as gui


def end_it(wb, excel):
    """kill excel
    """
    wb.Close(False)
    excel.Quit()
    del excel
    excel = None
    wb = None


dia = datetime.now().strftime('%d')
ano = datetime.now().year
mes = datetime.now().strftime('%m')
dir_path = os.path.dirname(os.path.realpath(__file__))
# not real spreadsheeting, only testing
arq = {dir_path+'\\PAINEL_1.xlsx': [1],
       dir_path+'\\PAINEL_2.xlsx': [1, 2]
       }  # this would originally open 8 different files and select 4 sheets each
# this would select the file based on the month and year
arq_esc = f'{dir_path}\\MOV_24_12.xlsx'
# but for testing and sharing reasons it is fixed at 2024 december

escuros = {dir_path+'\\query mov.xlsx': [1],
           arq_esc: []}

for file, _ in arq.items():
    # queries would update upon opening files
    os.system(f'start "excel" "{file}"')

time.sleep(20)
os.system("taskkill /f /im excel.exe")
time.sleep(2)

for file, index in arq.items():
    WB_PATH = file
    ws_index_list = index
    nome = file.split('_')[-1][:-5]
    dir_fim = f'{dir_path}\\{nome}'
    dir_dia = f'_{ano}_{mes}_{dia}'
    PATH_TO_PDF = dir_fim+dir_dia+'.pdf'
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        print('Start conversion to PDF')
        wb = excel.Workbooks.Open(WB_PATH)
        wb.Worksheets(ws_index_list).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
        time.sleep(5)
    except Exception as e:
        print('failed.')
        print(e)
        end_it(wb, excel)
        time.sleep(3)
    else:
        print('Succeeded.')
    finally:
        try:
            end_it(wb, excel)
            time.sleep(3)
        except Exception as e:
            print('failed.')
            print(e)
            os.system("taskkill /f /im excel.exe")
            time.sleep(3)

time.sleep(10)

input('press for autogui ')
for file, _ in escuros.items():
    if 'query' in file:
        os.system(f'start "excel" "{file}"')
        gui.getWindowsWithTitle("query mov")[0].maximize()
        time.sleep(10)
        gui.click(700, 120)
        gui.mouseDown(button='left')
        gui.moveTo(100, 120, duration=2)
        gui.mouseUp(button='left')
        gui.hotkey('ctrl', 'c')
    else:
        WB_PATH = file
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(WB_PATH)
        wb.Worksheets('test').Activate()
        time.sleep(3)
        gui.getWindowsWithTitle("MOV_24_12")[0].activate()
        time.sleep(1)
        gui.getWindowsWithTitle("MOV_24_12")[0].maximize()
        gui.click(200, 800)
        gui.scroll(10000)
        gui.click(100, 135)
        time.sleep(1)
        gui.hotkey('ctrl', 'v')
        time.sleep(3)
        gui.hotkey('printscreen')
        time.sleep(4)
        gui.click(1480, 670)
        time.sleep(1)
        gui.mouseDown(button='right')
        time.sleep(1)
        gui.moveTo(880, 300, duration=2)
        time.sleep(1)
        gui.mouseUp(button='right')
        time.sleep(1)
