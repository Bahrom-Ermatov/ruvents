import pandas as pd
import sympy.ntheory as nt
from datetime import datetime
import calendar

#Загружаем данные из файла
data = pd.read_excel ('task_support.xlsx', sheet_name="Tasks")

#Создаем list-ы для загрузки результата
res_num1 = []
res_num2 = []
res_num3 = []
res_date1 = []
res_date2 = []
res_date3 = []

#Перебираем элементы первой колонки(num1)
for i in range(1, len(data.num1)-1):
    if data.num1[i] % 2 == 0:   #Определяем четные числа
        res_num1.append(data.num1[i])

#Перебираем элементы второй колонки(num2)
for i in range(1, len(data.num2)-1):
    if nt.isprime(data.num2[i]):    #Определяем простые числа
        res_num2.append(data.num2[i])

#Перебираем элементы третьей колонки(num3)
for i in range(1, len(data.num3)-1):
    if float(data.num3[i].replace(' ', '').replace(',', '.')) < 0.5:    #Определяем числа, которые меньше 0.5
        res_num3.append(data.num3[i])

#Перебираем элементы четвертой колонки(date1)
for i in range(1, len(data.date1)-1):
    index = data.date1[i].find("Tue")
    if index >= 0:                      #Определяем вторники
        res_date1.append(data.date1[i])

#Перебираем элементы пятой колонки(date2)
for i in range(1, len(data.date2)-1):
    if datetime.strptime(data.date2[i], "%Y-%m-%d %H:%M:%S.%f").strftime("%A") == "Tuesday":        #Определяем вторники
        res_date2.append(data.date2[i])

#Перебираем элементы шестой колонки(date3)
for i in range(1, len(data.date3)-1):
    date = datetime.strptime(data.date3[i], "%m-%d-%Y")
    if date.strftime("%A") == "Tuesday":        #Определяем вторник
        day_of_month = date.strftime("%d")      #Определяем день месяца
        las_day_of_month = calendar.monthrange(int(date.strftime("%Y")), int(date.strftime("%m")))[1] #Определяем последний день месяца
        if int(las_day_of_month) - int(day_of_month) < 7:   #Затем определяем последний вторник
            res_date3.append(data.date3[i])

#Записываем итоговые данные в Dataframe
df_info = pd.DataFrame(
    {   'Итого': [
        'Количество четных чисел - ' + str(len(res_num1)),
        'Количество простых чисел - ' + str(len(res_num2)),
        'Количество чисел меньших 0.5 - ' + str(len(res_num3)),
        'Количество вторников - ' + str(len(res_date1)),
        'Количество вторников - ' + str(len(res_date2)),
        'Количество последних вторников - ' + str(len(res_date3))
        ]
    }
)

#Записываем детальные данные в Dataframe
df_num1 = pd.DataFrame({'num1': res_num1})
df_num2 = pd.DataFrame({'num2': res_num2})
df_num3 = pd.DataFrame({'num3': res_num3})
df_date1 = pd.DataFrame({'date1': res_date1})
df_date2 = pd.DataFrame({'date2': res_date2})
df_date3 = pd.DataFrame({'date3': res_date3})

#Записываем все данные в Excel
writer = pd.ExcelWriter('./result.xlsx', engine='xlsxwriter')  # pylint: disable=abstract-class-instantiated
df_info.to_excel(writer, sheet_name='summary', index=False)
df_num1.to_excel(writer, sheet_name='num1', index=False)
df_num2.to_excel(writer, sheet_name='num2', index=False)
df_num3.to_excel(writer, sheet_name='num3', index=False)
df_date1.to_excel(writer, sheet_name='date1', index=False)
df_date2.to_excel(writer, sheet_name='date2', index=False)
df_date3.to_excel(writer, sheet_name='date3', index=False)

writer.save()

