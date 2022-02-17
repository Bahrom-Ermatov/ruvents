import pandas
import sympy.ntheory as nt
from datetime import datetime
import calendar

#Загружаем данные из файла
data = pandas.read_excel ('task_support.xlsx', sheet_name="Tasks")

#Получаем колонки в виде списка
col_num1 = list(data.num1)
col_num2 = list(data.num2)
col_num3 = list(data.num3)
col_date1 = list(data.date1)
col_date2 = list(data.date2)
col_date3 = list(data.date3)

#Создаем list-ы для загрузки результата
res_num1 = []
res_num2 = []
res_num3 = []
res_date1 = []
res_date2 = []
res_date3 = []


#Перебираем элементы первой колонки(num1) в цикле
for i in range(1, len(col_num1)-1):
    if col_num1[i] % 2 == 0:
        res_num1.append(col_num1[i])

#Перебираем элементы второй колонки(num1) в цикле
for i in range(1, len(col_num2)-1):
    if nt.isprime(col_num2[i]):
        res_num2.append(col_num1[i])

#Перебираем элементы третьей колонки(num1) в цикле
for i in range(1, len(col_num3)-1):
    if float(col_num3[i].replace(' ', '').replace(',', '.')) < 0.5:
        res_num3.append(col_num3[i])

#Перебираем элементы четвертой колонки(date1) в цикле
for i in range(1, len(col_date1)-1):
    index = col_date1[i].find("Tue")
    if index >= 0:
        res_date1.append(col_date1[i])

#Перебираем элементы пятой колонки(date2) в цикле
for i in range(1, len(col_date2)-1):
    if datetime.strptime(col_date2[i], "%Y-%m-%d %H:%M:%S.%f").strftime("%A") == "Tuesday":
        res_date2.append(col_date2[i])

#Перебираем элементы шестой колонки(date3) в цикле
for i in range(1, len(col_date3)-1):
    date = datetime.strptime(col_date3[i], "%m-%d-%Y")
    if date.strftime("%A") == "Tuesday":
        day_of_month = date.strftime("%d")
        las_day_of_month = calendar.monthrange(int(date.strftime("%Y")), int(date.strftime("%m")))[1]
        if int(las_day_of_month) - int(day_of_month) < 7:
            res_date3.append(col_date3[i])

print(res_date3)




