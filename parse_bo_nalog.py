#!/usr/bin/env python
# coding: utf-8

# In[6]:

def main():
    import openpyxl
    import pandas as pd
    import numpy as np
    import requests
    from zipfile import ZipFile
    from io import BytesIO
    import json
    from datetime import datetime

    import tkinter as tk
    from tkinter import filedialog

    import warnings
    warnings.simplefilter("ignore")


    # In[2]:


    # запрашиваем эксель с ИНН (можно подавать по строкам, можно - через запятую в первой ячейке)

    root = tk.Tk()
    root.withdraw()

    path = filedialog.askopenfilename()


    # In[3]:


    # Забираем ИНН в список

    input_inns = pd.read_excel(path, header = None)

    if ',' in input_inns[0]:
        inns = [i.strip(' ') for i in input_inns[0].str.split(',')[0]]
    else: 
        inns = [i for i in input_inns.iloc[:,0]]


    # In[4]:


    # Находим id для каждого ИНН в базе данных на сайте, который будем использовать для дальнейшего поиска

    data_ids = {}

    for inn in inns:
        url = f'https://bo.nalog.ru/nbo/organizations/search?query={inn}&page=0'

        headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
        result = requests.get(url, headers=headers)
        data_id = result.json()['content'][0]['id']
        data_ids[str(data_id)] = inn


    # In[5]:


    # Отправляем запрос

    list_of_lists = []
    year_now = datetime.now().year

    for i in data_ids.keys():
        for year in range(2019, year_now):
            url_id = f'https://bo.nalog.ru/download/bfo/{i}?auditReport=false&balance=false&capitalChange=false&clarification=false&targetedFundsUsing=false&correctionNumber=0&financialResult=true&fundsMovement=false&type=XLS&period={year}'
            result = requests.get(url_id, headers=headers)
            
    # Загружаем зип, распоковываем, выбираем нужные данные
            
            if result.ok:
                zipfile = ZipFile(BytesIO(result.content))
                
                df1 = pd.read_excel(zipfile.open(f'{zipfile.infolist()[0].filename}'), sheet_name = [0], engine = 'openpyxl')
                df2 = pd.read_excel(zipfile.open(f'{zipfile.infolist()[0].filename}'), sheet_name = [1], skiprows = 4, engine = 'openpyxl')
                
                df1 = pd.concat(df1.values(),keys = df1.keys(), axis = 1).iloc[:,[0,6]]
                df2 = pd.concat(df2.values(),keys = df2.keys(), axis = 1).iloc[:,[3,9]]
                
                df2.iloc[:,1] = df2.iloc[:,1].replace({'\(':'-', '\)':'',' ':'','-':np.nan},regex=True)
                
                name = df1.iloc[4,1]
                inn = df1.iloc[8,1]
                address = df1.iloc[14,1]
                gross_income = df2.iloc[1,1]
                sales = df2.iloc[2,1]
                val = df2.iloc[3,1]
                net_income = df2.iloc[17,1]

                values_list = [name, address, inn, year, gross_income, sales, val, net_income]
                list_of_lists.append(values_list)
                
                print(f'{i} {year} ok')
            
            else:
                print(f'{i} {year} failed')

    # Создаем датафрейм
    columns = [
    'Полное наименование',
    'Юр. адрес',
    'ИНН',
    'Год',
    'Выручка, тыс. руб.',
    'Себестоимость продаж, тыс. руб.', 
    r'Валовая прибыль (убыток), тыс. руб.',
    r'Чистая прибыль (убыток), тыс. руб.'
    ]

    df = pd.DataFrame(list_of_lists, columns = columns)
    df = df.astype({
        'Год': int,
        'Выручка, тыс. руб.': float,
        'Себестоимость продаж, тыс. руб.': float,
        'Валовая прибыль (убыток), тыс. руб.': float,
        'Чистая прибыль (убыток), тыс. руб.': float
    })

    df.index += 1


    # In[ ]:


    # Выгружаем датафрейм

    now = datetime.now().strftime("%d.%m.%Y_%H-%M-%S")
    new_path = '/'.join(path.split('/')[:-1])+f'/result_{now}.xlsx'
    df.to_excel(new_path)

if __name__ == '__main__':
    main()