from win32com.client.gencache import EnsureDispatch
import os
import re
import sqlite3
from pprint import pprint
import pandas as pd
import itertools
from itertools import groupby



# Get the Excel Application COM object
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\1.xlsx")
Sheets = wb.Sheets.Count
ws = wb.Worksheets(Sheets)

def main():
    L_crews, L_plates = [], []
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
            L_plates.append('*****')
        elif 'Вспомогательная техника' in str(ws.Cells(row, 4).Value):
            L_plates.append('*****')
        else:
            L_plates.append(ws.Cells(row, 6).Value)
        row += 1

        if 'Технолог ЦИТС' in str(ws.Cells(row, 6).Value):
            break
    
    # Cleaning main trucks
    # Remove unnecessary items
    L_remItems = ['Бурильщик', 'Пом.бур', 'профессия', ',']
    for i in L_remItems:
        for j in L_plates:
            if i == j:
                L_plates.remove(j)
              
    L_plates = [x for x in L_plates if x != None]
    L_plates = [x for x in L_plates if x != 'Ф.И.О.']
    
    # Cleaning necessary items 
    pattern = re.compile(r"(Маш-т)|(гос№)|(гос.№)|(Гос№)|(НТ)|(НКА)|(ПКА)|(МЗКТ)|(УГА)|(МАК)|(МК30Т)|(МАН)|(замена ДВС)|(ГНКТ)")
    L_plates = [re.sub(pattern, '', x).strip() for x in L_plates]
     
    L_plates = [str(x).split('\n') for x in L_plates]
    L_plates = list(itertools.chain.from_iterable(L_plates))
    
    L_plates = [str(x).split(',') for x in L_plates]
    L_plates = list(itertools.chain.from_iterable(L_plates))
    L_plates = [str(x).split('RUS') for x in L_plates]
    L_plates = list(itertools.chain.from_iterable(L_plates))
    
    # Cleaning trash remainder
    for i in L_plates:
        if 'ГТ' in i:
            L_plates.remove(i)
        else:
            L_plates = [x.strip() for x in L_plates if x != 'б/н']
    
    L_plates = [x for x in L_plates if x != '']
    
    pattern = re.compile(r"[)(-.№]")

    # Stripping regions
    L_plates = [re.sub(pattern, '', x).strip() for x in L_plates]
    L_plates = [x.removesuffix('186').strip() for x in L_plates]
    L_plates = [x.removesuffix('86').strip() for x in L_plates]
    L_plates = [x.removesuffix('54').strip() for x in L_plates]
    L_plates = [x.removesuffix('82').strip() for x in L_plates]
    L_plates = [x.removesuffix('77').strip() for x in L_plates]
    
    # Cleaning trash remainder
    L_plates = [x.replace(' ', '') for x in L_plates]
    L_plates = [x for x in L_plates if len(x) >= 3 or len(x) == 0]
    
    
    # Collecting multiplitaction indeces to match up with crews
    
    L_counts = []
    counter = 0
        
    for i in L_plates:
        if i != '':
            counter = 1
        else:
            counter = 0
        L_counts.append(counter)
    
    # Summing up ones between zeros
    L_counts = [i for k, g in groupby(L_counts, bool) for i in ((sum(g),) if k else g)]  
    L_counts = [x for x in L_counts if x != 0]
    
    # Multiplying crews to counts
    L_crews = [(i * j).split('ГНКТ') for i, j in (zip(L_counts, L_crews))]
    L_crews = list(itertools.chain.from_iterable(L_crews))
    L_crews = [x.replace('№', 'ГНКТ').strip() for x in L_crews if x != '']
    
    # Peeling Nones to match up with crews per length
    L_plates = [x.lower() for x in L_plates if x != '']
    
    # Building data frame
    data = pd.DataFrame(zip(L_crews, L_plates), columns = ['Crews', 'Plates'])
    

    # Making colum of plate indeces
    L_index = [''.join(re.findall(r'\d+', x)) for x in  L_plates]
    pprint(L_index)
    pprint(len(L_index))
    # Making colum of plate literals
    L_literal = [''.join(re.findall(r'[А-я]', x)) for x in L_plates]
    # print(L_literal)
    
    # Extracting table columns as lists from final_DB
    db = sqlite3.connect('omnicomm.db')
    db.row_factory = lambda cursor, row: row[0]
    cursor = db.cursor()

    L_omn_vehicle = cursor.execute("SELECT Vehicle_Name FROM final_DB").fetchall()
    L_omn_plate = cursor.execute("SELECT Plate FROM final_DB").fetchall()
    L_omn_index = cursor.execute("SELECT Plate_index FROM final_DB").fetchall()
    L_omn_literal = cursor.execute("SELECT Plate_literal FROM final_DB").fetchall()

    # find plates in omnicom by index and literal, build list of vehicle names
    # Forming temp list
    L_ct_vehicle = []
    L_ct_plate = []
    for i, j in zip(L_index, L_literal):
        if i in L_omn_index and j in L_omn_literal:
            ind = L_omn_index.index(i)
            L_ct_vehicle.append(L_omn_vehicle[ind])
            L_ct_plate.append(L_omn_plate[ind])
        else:
            L_ct_vehicle.append('N/A')
            L_ct_plate.append('N/A')

    data = pd.DataFrame(zip(L_crews, L_ct_vehicle, L_ct_plate), columns = ['Бригада', 'СПТ', 'Гос.номер'])
    print(data)
    # print(len(L_literal))
    # print(len(L_index))
    # print(len(L_veh))
    # Posting data frame to DB

    # L_plates = [str(x).replace(';', ',').split(',') for x in L_plates if x != None]

    # L_replace = ['+', '(ПС)', '86-', ')']
    
    # for i in L_replace:
    #     L_plates = [''.join(str(x).split(str(i))) for x in L_plates]    
    # L_plates = list(itertools.chain.from_iterable(L_plates))
    # L_plates = [str(x).split('+') for x in L_plates]
    # L_plates = list(itertools.chain.from_iterable(L_plates))
    # L_plates = [str(x).split('(ПС)') for x in L_plates]
    # L_plates = list(itertools.chain.from_iterable(L_plates))
    # L_plates = [str(x).split('86-') for x in L_plates]
    # L_plates = list(itertools.chain.from_iterable(L_plates))
    # L_plates = [str(x).split(')') for x in L_plates]
    # L_plates = list(itertools.chain.from_iterable(L_plates))
    # L_plates = [str(x).strip() for x in L_plates]
    # print(len(L_plates))
    # L_plates = [x for x in L_plates if re.findall("\d+", x) or x == '*****']
    # L_stars = [x for x in L_plates if x == '*****']

    # L_remItems = ['Желобная', 'ёмк', 'Жел.', 'гуммиров', 'гуммированная ёмкость', 'Гум.', 'Гум.емкость', 'Инструм', 'Столовая', 'Спаль','Спальный', 'транспортный', 
    #             'Нива', 'НИВА', 'Вагон', 'Мастера', 'мастера', 'ДЭС', 'ДЭС-', 'контейнер', 'Сушил', 'Пуст.кат', 'жил.', 'резервная']

    # for i in L_remItems:
    #     for j in L_plates:
    #         if re.search(i, str(j)):
    #             L_plates.remove(j)

    
    # pattern = re.compile(r"(Маш-т)|(гос№)|(гос.№)|(Гос№)")
    # L_plates = [re.sub(pattern, '', x).strip() for x in L_plates]
    # pattern = re.compile(r"\(\d+")
    # L_plates = [re.sub(pattern, '', x).strip() for x in L_plates]
    # L_plates = [x.replace('\n', ' ') for x in L_plates]
    # L_plates = [x.replace('трал', 'трал-') for x in L_plates]
        
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
    #             for x in L_plates]

    # L_plates = [x.replace('186', '') for x in L_plates]
    
    # data = pd.DataFrame(zip(L_plates, L_plates))
    # pd.set_option("display.max_rows", None, "display.max_columns", None)
    # print(data)
    # pprint(L_plates)
    # print(len(L_plates))
    # pprint(L_plates)
    # pprint(len(L_plates))
    # print(L_crews)
    # print(len(L_crews))
    # print(len(L_stars))

if __name__ == '__main__':
    main()


