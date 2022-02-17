import pandas
from geopy import distance
import itertools
import csv
import datetime

now = datetime.datetime.now()

data = pandas.read_excel ('short_list.xlsx', sheet_name="Лист1", usecols=['Конечный населенный пункт', 'lon', 'lat'])

settlements = data.to_dict(orient='record')

settlement_names = ['#']
for row in settlements:
    settlement_names.append(row['Конечный населенный пункт'])

csvfile=open('Distance.csv','w', newline='')
writer=csv.writer(csvfile, delimiter=';')
writer.writerows([settlement_names])

data = []
length = len(settlements)

for i in range(length): 
    line = [settlements[i]['Конечный населенный пункт']]
    for j in range(length): 
        if (i == j):
            line.append('0км')
        elif (i > j):
            line.append(data[j][i+1])
        else:
            line.append(str(round(distance.distance((settlements[i]["lat"], settlements[i]["lon"]), (settlements[j]["lat"], settlements[j]["lon"])).km, 3))+"км")

    writer.writerows([line])
    data.append(line)

csvfile.close()

print("Время выполнения = ", datetime.datetime.now() - now)
