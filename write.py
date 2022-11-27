import os

import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
from excel_read import Excel2Dic, path_out

Aluminium_Production_Asia = Excel2Dic('Aluminium_Production_Asia')
Cement_Production_Asia = Excel2Dic('Cement_Production_Asia')
Copper_Production_Asia = Excel2Dic('Copper_Production_Asia')
GDP_Asia = Excel2Dic('GDP_Asia')
Paper_Production_Asia = Excel2Dic('Paper_Production_Asia')
People_Asia = Excel2Dic('People_Asia')
Steel_Production_Asia = Excel2Dic('Steel_Production_Asia')
xls_collection ={}
value = [Aluminium_Production_Asia, Cement_Production_Asia, Copper_Production_Asia, GDP_Asia,
                  Paper_Production_Asia, People_Asia, Steel_Production_Asia]
xls_name = ['Aluminium_Production_Asia', 'Cement_Production_Asia', 'Copper_Production_Asia', 'GDP_Asia',
            'Paper_Production_Asia', 'People_Asia', 'Steel_Production_Asia']
for k,v in zip(xls_name,value):
    xls_collection.setdefault(k,v)

a =GDP_Asia.get_value('1960','Afghanistan')


def search_value(xls_dic, country):
    key = []
    for forecast_year in range(1960, 2022):
        value = xls_dic.get_value(str(forecast_year), country)
        key.append(value)

    return np.array(key)

def draw(x,y,country_name):
    # label =list(str(i)for i in range(1960,2022))
    # plt.xticks( people,label,rotation ='vertical')
    plt.scatter(x, y * 10000)
    plt.title(country_name)
    plt.show()

def start(country):
    column_names_aps = []
    for forecast_year in range(1960, 2022):
        column_names_aps.append(forecast_year)

    people = search_value(xls_collection['People_Asia'], country)

    Aluminium = search_value(xls_collection['Aluminium_Production_Asia'], country)
    Aluminium_people = Aluminium / people

    Cement = search_value(xls_collection['Cement_Production_Asia'], country)
    Cement_people = Cement / people

    Copper = search_value(xls_collection['Copper_Production_Asia'], country)
    Copper_people = Copper / people

    GDP = search_value(xls_collection['GDP_Asia'], country)

    Paper = search_value(xls_collection['Paper_Production_Asia'], country)
    Paper_people = Paper / people

    Steel = search_value(xls_collection['Steel_Production_Asia'], country)
    Steel_people = Steel / people

    col_content_aps = [column_names_aps, people, GDP, Aluminium, Aluminium_people, Cement, Cement_people,
                       Copper, Copper_people, Paper, Paper_people, Steel, Steel_people]
    result = pd.DataFrame(col_content_aps)
    result.insert(0, 'Year', ['Year', 'People', 'GDP per capita (constant 2017 international $)',
                                     'Aluminium production (thousand short ton)', 'Aluminium production per capita',
                                     'Cement production (thousand short ton)','Cement production per capita',
                                     'Copper production (thousand short ton)','Copper production per capita',
                                     'Paper production (thousand short ton)','Paper productionper capita',
                                     'Steel production (thousand short ton)', 'Steel  productionper capita'])



    # draw(GDP,Paper_people,country_name=country)
    # draw(GDP, Aluminium_people, country_name=country+'AL')

    return result

outfile=os.path.join(path_out, f'sum.xlsx')
with pd.ExcelWriter(outfile) as writer:

    for country in People_Asia.contoury_name:
        result = start(country)
        result.to_excel(writer,sheet_name=f'{country}', index=False, header=False)


