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

    row = 4
    while True:
        if 'ГНКТ №' in str(ws.Cells(row, 1).Value):
            L_crews.append(ws.Cells(row, 1).Value)
        if 'Технолог ЦИТС' in str(ws.Cells(row, 6).Value):
            break   
        row += 1

    row = 4
    while True:
        # if 'ГНКТ' in str(ws.Cells(row, 1).Value):
        #     L_crews.append(ws.Cells(row, 1).Value)
        if 'Цель работ:' in str(ws.Cells(row, 6).Value):
            L_vehicles.append('*****')
        else:
            L_vehicles.append(ws.Cells(row, 6).Value)
        row += 1

        if 'Технолог ЦИТС' in str(ws.Cells(row, 6).Value):
            break
    L_vehicles = [str(x).replace(';', ',').split(',') for x in L_vehicles if x != None]
    L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    L_vehicles = [str(x).split('+') for x in L_vehicles]
    L_vehicles = list(itertools.chain.from_iterable(L_vehicles))
    L_vehicles = [str(x).strip() for x in L_vehicles]
    print(len(L_vehicles))
    L_vehicles = [x for x in L_vehicles if re.findall("\d+", x) or x == '*****']
    L_stars = [x for x in L_vehicles if x == '*****']

    L_remItems = ['Желобная', 'ёмк', 'Жел.', 'гуммиров', 'гуммированная ёмкость', 'Гум.', 'Гум.емкость', 'Инструм', 'Столовая', 'Спаль','Спальный', 'транспортный', 
                'Нива', 'НИВА', 'Вагон', 'Мастера', 'мастера', 'ДЭС', 'ДЭС-', 'контейнер', 'Сушил', 'Пуст.кат', 'жил.', 'резервная']

    for i in L_remItems:
        for j in L_vehicles:
            if re.search(i, str(j)):
                L_vehicles.remove(j)

    
    pattern = re.compile(r"(Маш-т)|(гос№)|(гос.№)|(Гос№)")
    L_vehicles = [re.sub(pattern, '', x).strip() for x in L_vehicles]
    # changing this С/Т-845(694ПС) to С/Т-845(ПС)
    # pattern = re.compile(r"((\(\d+"))|(\(\d+\W+))")
    pattern = re.compile(r"\(\d+")
    L_vehicles = [re.sub(pattern, '', x).strip() for x in L_vehicles]

    
    L_vehicles = [x.replace('\n', ' ') for x in L_vehicles]
    L_vehicles = [x.replace('трал', 'трал-') for x in L_vehicles]
        
            
    plates1 = re.compile("[А-Яа-я]*\d+[А-Яа-я]{2}\s*\d+")
    plates2 = re.compile("[А-Яа-я]{2}\d+\s\d+")
    plates3 = re.compile("\w{2}\s\D\d+\s\d{2}")
    plates4 = re.compile("\d+\s*[А-Яа-я]{2}\s*\d+")
    plates5 = re.compile("[А-Яа-я]{2}\s*\d+\s*\d+")
    plates6 = re.compile("\d")
    plates7 = re.compile("[А-Яа-я]\s*\d+\s*\W+\s*\d+")

    L_plates = [''.join(re.findall(plates7, x)) or
                ''.join(re.findall(plates1, x)) or 
                ''.join(re.findall(plates2, x)) or 
                ''.join(re.findall(plates3, x)) or 
                ''.join(re.findall(plates4, x)) or 
                ''.join(re.findall(plates5, x)) or
                ''.join(re.findall(plates6, x)) 
                
                for x in L_vehicles]
    
    data = pd.DataFrame(zip(L_vehicles, L_plates))
    pd.set_option("display.max_rows", None, "display.max_columns", None)
    print(data)
    # pprint(L_vehicles)
    # print(len(L_vehicles))
    # pprint(L_plates)
    # pprint(len(L_plates))
    # print(L_crews)
    # print(len(L_crews))
    # print(len(L_stars))

if __name__ == '__main__':
    main()


