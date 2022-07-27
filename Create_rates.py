import openpyxl
from openpyxl import *
import pandas as pd

paths = ['2022_data', '2021_data', '2020_data', '2019_data']
for path in paths:
    workbook = openpyxl.load_workbook(f'C:/Users/ken/Desktop/Elula_Project/{path}.xlsx', read_only=True, data_only=True)
    print("workbook has been opened")

    # contains all sheetnames)
    sheetnames = workbook.sheetnames
    workbook.close()

    dict = {}
    xlsx_path = f'C:/Users/ken/Desktop/Elula_Project/{path}.xlsx'
    xlsx_file = pd.ExcelFile(f'C:/Users/ken/Desktop/Elula_Project/{path}.xlsx')
    for i in range(0, len(sheetnames)):
        key = f'df{i + 1}'
        value = pd.read_excel(xlsx_file, sheetnames[i])
        dict[key] = value

    for key, value in dict.items():
        dict[key].columns = dict[key].iloc[0]
        df = dict[key].iloc[1:].reset_index(drop=True)
        df = df.drop(index=[0, 1, 7, 8, 19, 21, 30, 31, 33, 36, 43, 48, 49, 50, 54, 55], axis=0)
        # df2 = df.drop('H', axis=1).reset_index(drop=True)
        df2 = df.drop(['H', 'Mortgage interest rates'], axis=1).reset_index(drop=True)
        df2.to_csv(f'data{key}.csv', index=False)

    final_data = []
    for key, value in dict.items():
        df = pd.read_csv(f'data{key}.csv')
        df = df.loc[:, ~df.columns.str.contains('^Unnamed: 8')]
        df.rename(columns={'Unnamed: 0': '1', 'Unnamed: 1': '2', 'Unnamed: 2': '3', 'Unnamed: 3': '4', 'Unnamed: 4': '5',
                           'Unnamed: 5': '6', 'Unnamed: 6': '7', 'Unnamed: 7': '8'}, inplace=True)
        df_list = df.values.tolist()

        for list in df_list:
            for x in range(len(list)):
                final_data.append(list[x])

    print(final_data)
    final_dataframe = pd.DataFrame(final_data, columns=['Rate'])
    print(final_dataframe)
    final_dataframe.to_csv(f'Final_{path}_data.csv')

