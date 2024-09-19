import pandas as pd
import xlwt
import os

# Парсинг файла и создание словаря прогульщиков
def progul_list(filename):
    men = dict()
    with open(f'{os.getcwd()}\{filename}.txt', 'r', encoding='utf-8') as fol:
        file = fol.read().split('\n')
        for i in file:
            if i !='' or i.split(' ')[0].isalpha():
                try:
                    if i.split(' ')[0] not in men.keys():
                        
                            men[i.split(' ')[0]] = (i.split(' ')[1])
                    else:
                        men[i.split(' ')[0]] = men[i.split(' ')[0]]+(i.split(' ')[1])
                except:
                        pass
                print(i)
        print(len(file))

    s = {i : '+'.join(men[i]) for i in men.keys()}

    itog_n = {}
    for i in s.keys():
        try: 
            itog_n[i] = sum([int(i) for i in s[i].split('+')])
        except: pass

    itog = {j : itog_n[j]*2 for j in sorted([i for i in itog_n])}
    pr = ['Из них прогулов']
    pr+=[itog[i] for i in itog]
    return pr




row_index = 1
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet1")

df = pd.read_excel('nb.xlsx', index_col='№')

df = df.rename(columns={'Ф.И.О':'ФИО'})

df.columns

how_much_nb = ['Всего пропусков']
# print(df['ФИО'].values)
with open('nb_f.txt', 'w', encoding='utf-8') as f: f.write('\n')

names = ['ФИО студентов']
for i in df['ФИО'].values: # Чтение данных из таблицы
    
    names.append(i)

    df_temp = df[df['ФИО'] == i].T.dropna()

    df_temp.drop('ФИО', inplace=True)

    df_temp.index = df_temp.index.astype('float64')

    df_temp.sort_index(inplace=True)

    df_temp.reset_index(inplace=True)

    df_temp.columns = ['date','nb']

    progul = {i:' ' for i in range(1, 32)}
    print(df_temp)
    props = 0
    for i in df_temp['nb']: # Тут происходит обнаружение "нб" и перевод их в часы пропусков
        if i == 'нб':
            props+=2
    # print([i for i in df_temp['nb']])
    how_much_nb.append(props)
    for k, j in enumerate(df_temp.values): # Создание словаря
        print(df_temp.values[k])
        day = int(str(df_temp.values[k]).split('.')[0][1::])
        print(day)
        progul[day] =  'нб' # Занесение значения к ключу словаря

    kolvo = "\n".join("{!r} - {!r}".format(k, v) for k, v in progul.items() if v != 0)
    with open('nb_f.txt', 'a', encoding='utf-8') as f: f.write(f"{'_'*30}\n{i}\n{kolvo}\n{len(df_temp['nb'])*2}\n{'-'*30}\n\n")

    for j in range(1, len(progul)+1): # Добавление нб по ячейкам
        print(j, progul[j])
        row = sheet1.row(row_index)
        value = progul[j]
        row.write(j, value)
    row_index += 1

for i, j in enumerate(progul): # Добавление дат
    row = sheet1.row(0)
    row.write(i+1, j)

for i, j in enumerate(names): # Добавление ФИО
    row = sheet1.row(i)
    row.write(0, j)

for i, j in enumerate(how_much_nb): # Добавление количества пропусков
    row = sheet1.row(i)
    row.write(32, j)

for i, j in enumerate(progul_list('nb_f')): # Добавление количества прогулов
    row = sheet1.row(i)
    print(j)
    row.write(33, j)

book.save("NB table.xls")


# print(progul[1])
# print([f'{i} - {progul[i]}' for i in progul if progul[i] != 0])
# print('234567')
# print("{" + ",\n".join("{!r}: {!r}".format(k, v) for k, v in progul.items() if v != 0) + "}")


