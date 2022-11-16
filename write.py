import os
import numpy as np
import pandas as pd
from excel_read import Excel2Dic, path_out

Aluminium_Production_Asia = Excel2Dic('Aluminium_Production_Asia')
Cement_Production_Asia = Excel2Dic('Cement_Production_Asia')
Copper_Production_Asia = Excel2Dic('Copper_Production_Asia')
GDP_Asia = Excel2Dic('GDP_Asia')
Paper_Production_Asia = Excel2Dic('Paper_Production_Asia')
People_Asia = Excel2Dic('People_Asia')
Steel_Production_Asia = Excel2Dic('Steel_Production_Asia')

xls_collection = [Aluminium_Production_Asia, Cement_Production_Asia, Copper_Production_Asia, GDP_Asia,
                  Paper_Production_Asia, People_Asia, Steel_Production_Asia]
xls_name = ['Aluminium_Production_Asia', 'Cement_Production_Asia', 'Copper_Production_Asia', 'GDP_Asia',
            'Paper_Production_Asia', 'People_Asia', 'Steel_Production_Asia']


def search_value(xls_dic, country):
    key = []
    for forecast_year in range(1960, 2022):
        value = xls_dic.get_value(str(forecast_year), country)
        key.append(value)

    return np.array(key)


def start(country):
    column_names_aps = []
    for forecast_year in range(1960, 2022):
        column_names_aps.append(forecast_year)

    people = search_value(xls_collection[5], country)

    alum = search_value(xls_collection[0], country)
    alum_people = alum / people

    Cement = search_value(xls_collection[1], country)
    Cement_people = Cement / people

    Copper = search_value(xls_collection[2], country)
    Copper_people = Copper / people

    GDP = search_value(xls_collection[3], country)

    Paper = search_value(xls_collection[4], country)
    Paper_people = Paper / people

    Steel = search_value(xls_collection[6], country)
    Steel_people = Steel / people

    col_content_aps = [column_names_aps, people, GDP, alum, alum_people, Cement, Cement_people,
                       Copper, Copper_people, Paper, Paper_people, Steel, Steel_people]
    prognosis_aps = pd.DataFrame(col_content_aps)

    prognosis_aps.insert(0, 'Year', ['Year', 'People', 'GDP', 'alum',
                                     'alum_people', 'Cement',
                                     'Cement_people', 'Copper',
                                     'Copper_people', 'Paper',
                                     'Paper_people', 'Steel',
                                     'Steel_people'])

    return prognosis_aps
outfile=os.path.join(path_out, f'sum.xlsx')
with pd.ExcelWriter(outfile) as writer:

    for country in People_Asia.contoury_name:
        result = start(country)
        result.to_excel(writer,sheet_name=f'{country}', index=False, header=False)


