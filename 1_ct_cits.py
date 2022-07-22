from win32com.client.gencache import EnsureDispatch
import os
import re
from pprint import pprint
import pandas as pd
import itertools



# Get the Excel Application COM object
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\1.xlsx")
Sheets = wb.Sheets.Count
ws = wb.Worksheets(Sheets)

def main():
    L_crews, L_vehicles = [], []
    # Adding crews to array
    row = 4
    while True:
        if 'ГНКТ №' in str(ws.Cells(row, 1).Value):
            L_crews.append(ws.Cells(row, 1).Value)
        if 'Технолог ЦИТС' in str(ws.Cells(row, 6).Value):
            break   
        row += 1
    # Adding main trucks to array
    row = 4
    while True:
        
        if 'Цель работ:' in str(ws.Cells(row, 6).Value):
            L_vehicles.append('*****')
        elif 'Вспомогательная техника' in str(ws.Cells(row, 4).Value):
            L_vehicles.append('*****')
        else:
            L_vehicles.append(ws.Cells(row, 6).Value)
        row += 1

        if 'Технолог ЦИТС' in str(ws.Cells(row, 6).Value):
            break
    # Cleaning main trucks
    # Remove unnecessary items
    L_remItems = ['Бурильщик', 'Пом.бур', 'профессия', ',']
    for i in L_remItems:
        for j in L_vehicles:
            if i == j:
                L_vehicles.remove(j)
                
    L_vehicles = [x for x in L_vehicles if x != None]
    L_vehicles = [x for x in L_vehicles if x != 'Ф.И.О.']
    # Cleaning necessary items to plates
    pattern = re.compile(r"(Маш-т)|(гос№)|(гос.№)|(Гос№)|(НТ)|(НКА)|(ПКА)|(МЗКТ)|(УГА)|(МАК)|(МК30Т)|(МАН)|(замена ДВС)|(ГНКТ)")
    L_vehicles = [re.sub(pattern, '', x).strip() for x in L_vehicles]
    L_vehicles = [str(x).split('\n') for x in L_vehicles]
    L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    L_vehicles = [str(x).split(',') for x in L_vehicles]
    L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    L_vehicles = [str(x).split('RUS') for x in L_vehicles]
    L_vehicles = list(itertools.chain.from_iterable(L_vehicles))

    for i in L_vehicles:
        if 'ГТ' in i:
            L_vehicles.remove(i)
        else:
            L_vehicles = [x.strip() for x in L_vehicles if x != 'б/н']

    pattern = re.compile(r"[)(-.№]")
    L_vehicles = [re.sub(pattern, '', x).strip() for x in L_vehicles]
    L_vehicles = [x.removesuffix('186').strip() for x in L_vehicles]
    L_vehicles = [x.removesuffix('86').strip() for x in L_vehicles]
    L_vehicles = [x.removesuffix('54').strip() for x in L_vehicles]
    L_vehicles = [x.removesuffix('82').strip() for x in L_vehicles]
    L_vehicles = [x.removesuffix('77').strip() for x in L_vehicles]
    
   
    L_vehicles = [x.replace(' ', '') for x in L_vehicles]
    L_vehicles = [x for x in L_vehicles if len(x) >= 3 or len(x) == 0]
    
    pprint(L_vehicles)
    
    # Collecting multiplitaction indeces to match up with crews

    # L_vehicles = [str(x).replace(';', ',').split(',') for x in L_vehicles if x != None]

    # L_replace = ['+', '(ПС)', '86-', ')']
    
    # for i in L_replace:
    #     L_vehicles = [''.join(str(x).split(str(i))) for x in L_vehicles]    
    # L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    # L_vehicles = [str(x).split('+') for x in L_vehicles]
    # L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    # L_vehicles = [str(x).split('(ПС)') for x in L_vehicles]
    # L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    # L_vehicles = [str(x).split('86-') for x in L_vehicles]
    # L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    # L_vehicles = [str(x).split(')') for x in L_vehicles]
    # L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    # L_vehicles = [str(x).strip() for x in L_vehicles]
    # print(len(L_vehicles))
    # L_vehicles = [x for x in L_vehicles if re.findall("\d+", x) or x == '*****']
    # L_stars = [x for x in L_vehicles if x == '*****']

    # L_remItems = ['Желобная', 'ёмк', 'Жел.', 'гуммиров', 'гуммированная ёмкость', 'Гум.', 'Гум.емкость', 'Инструм', 'Столовая', 'Спаль','Спальный', 'транспортный', 
    #             'Нива', 'НИВА', 'Вагон', 'Мастера', 'мастера', 'ДЭС', 'ДЭС-', 'контейнер', 'Сушил', 'Пуст.кат', 'жил.', 'резервная']

    # for i in L_remItems:
    #     for j in L_vehicles:
    #         if re.search(i, str(j)):
    #             L_vehicles.remove(j)

    
    # pattern = re.compile(r"(Маш-т)|(гос№)|(гос.№)|(Гос№)")
    # L_vehicles = [re.sub(pattern, '', x).strip() for x in L_vehicles]
    # pattern = re.compile(r"\(\d+")
    # L_vehicles = [re.sub(pattern, '', x).strip() for x in L_vehicles]
    # L_vehicles = [x.replace('\n', ' ') for x in L_vehicles]
    # L_vehicles = [x.replace('трал', 'трал-') for x in L_vehicles]
        
    # plates1 = re.compile("(\d{4})")    
    # plates2 = re.compile("(\d{3})")    
    # plates1 = re.compile("[А-Яа-я]*\d+[А-Яа-я]{2}\s*\d+")
    # plates2 = re.compile("[А-Яа-я]{2}\d+\s\d+")
    # plates3 = re.compile("\w{2}\s\D\d+\s\d{2}")
    # plates4 = re.compile("\d+\s*[А-Яа-я]{2}\s*\d+")
    # plates5 = re.compile("[А-Яа-я]{2}\s*\d+\s*\d+")
    # plates6 = re.compile("\d")
    # plates7 = re.compile("[А-Яа-я]\s*\d+\s*\W+\s*\d+")

    # L_plates = [''.join(re.findall(plates1, x)) or 
    #             ''.join(re.findall(plates2, x)) 
    #             for x in L_vehicles]

    # L_plates = [x.replace('186', '') for x in L_plates]
    
    # data = pd.DataFrame(zip(L_vehicles, L_plates))
    # pd.set_option("display.max_rows", None, "display.max_columns", None)
    # print(data)
    # pprint(L_vehicles)
    # print(len(L_vehicles))
    # pprint(L_plates)
    # pprint(len(L_plates))
    # print(L_crews)
    # print(len(L_crews))
    # print(len(L_stars))

if __name__ == '__main__':
    main()


